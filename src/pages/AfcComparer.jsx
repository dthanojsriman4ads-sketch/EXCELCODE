import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileText, Layout, Download, ArrowLeft, CheckCircle2, AlertCircle, TrendingUp, DollarSign, X } from 'lucide-react';
import { Link } from 'react-router-dom';

const AfcComparer = () => {
  const [reports, setReports] = useState({ orders: null, webhook: null, afc: null });
  const [filenames, setFilenames] = useState({ orders: '', webhook: '', afc: '' });
  const [afcFilter, setAfcFilter] = useState('DTVM'); // 'DTVM' or 'NOT_DTVM'
  const [results, setResults] = useState(null);
  const [isUploading, setIsUploading] = useState({ orders: false, webhook: false, afc: false });
  const [isProcessing, setIsProcessing] = useState(false);

  const handleFileUpload = (e, type) => {
    const file = e.target.files[0];
    if (!file) return;

    setIsUploading(prev => ({ ...prev, [type]: true }));
    setFilenames(prev => ({ ...prev, [type]: file.name }));

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
        setReports(prev => ({ ...prev, [type]: data }));
      } catch (err) {
        console.error("File upload error:", err);
        alert("Failed to read file. Please ensure it is a valid Excel file.");
      } finally {
        setIsUploading(prev => ({ ...prev, [type]: false }));
      }
    };
    reader.onerror = () => {
      alert("Error reading file.");
      setIsUploading(prev => ({ ...prev, [type]: false }));
    };
    reader.readAsBinaryString(file);
  };

  const detectHeaders = (data, keywords) => {
    let mapping = { headerIndex: -1 };
    for (let i = 0; i < Math.min(data.length, 20); i++) {
      const row = (data[i] || []).map(c => c?.toString().toLowerCase().trim() || "");
      if (row.length === 0) continue;

      let matchCount = 0;
      let tempMapping = {};

      row.forEach((cell, idx) => {
        Object.entries(keywords).forEach(([key, matches]) => {
          // Check for exact match first, then partial
          if (matches.some(m => cell === m || cell.includes(m))) {
            if (tempMapping[key] === undefined || cell === keywords[key][0]) {
              tempMapping[key] = idx;
            }
          }
        });
      });

      // We consider it a header row if at least ID and Revenue are found
      if (tempMapping.id !== undefined && tempMapping.revenue !== undefined) {
        return { ...tempMapping, headerIndex: i };
      }
    }
    return mapping;
  };

  const parseCurrency = (val) => {
    if (val === undefined || val === null || val === '') return 0;
    const clean = val.toString().replace(/[^\d.-]/g, '');
    const num = parseFloat(clean);
    return isNaN(num) ? 0 : num;
  };

  const processComparison = () => {
    if (!reports.orders && !reports.webhook) {
      alert("Please upload at least one order report (Orders or Webhook).");
      return;
    }
    if (!reports.afc) {
      alert("Please upload the AFC report.");
      return;
    }

    setIsProcessing(true);
    setResults(null);

    // Using setTimeout to allow the UI to render the "Processing" state
    setTimeout(() => {
      try {
        const orderKeywords = {
          id: ['pg_order_id', 'pg_order', 'order_id', 'orderid'],
          revenue: ['total_amount', 'total amount', 'fare', 'amount'],
          type: ['ticket_type', 'ticket type', 'type'],
          tickets: ['num_of_tickets', 'no of tickets', 'tickets', 'count']
        };

        const afcKeywords = {
          id: ['merchant_order_id', 'merchant_order', 'merchant order', 'order_id', 'orderid'],
          revenue: ['actual_fare', 'actual fare', 'fare', 'amount'],
          type: ['ticket_type', 'ticket type', 'type'],
          status: ['ticket_status', 'ticket status', 'status']
        };

        const ordersMap = reports.orders ? detectHeaders(reports.orders, orderKeywords) : { headerIndex: -1 };
        const webhookMap = reports.webhook ? detectHeaders(reports.webhook, orderKeywords) : { headerIndex: -1 };
        const afcMap = detectHeaders(reports.afc, afcKeywords);

        if (afcMap.headerIndex === -1) {
          alert("Could not detect headers in AFC file.");
          setIsProcessing(false);
          return;
        }

        const combinedOrdersMap = new Map();
        let ourTotalRevenue = 0;

        const processRows = (data, mapping) => {
          if (mapping.headerIndex === -1) return;
          const rows = data.slice(mapping.headerIndex + 1);
          rows.forEach(row => {
            const id = row[mapping.id]?.toString().trim();
            if (!id) return;
            
            const rev = parseCurrency(row[mapping.revenue]);
            const type = row[mapping.type]?.toString().trim().toLowerCase() || 'single';
            const tickets = parseInt(row[mapping.tickets]) || 1;
            
            const expectedCount = type.includes('return') ? tickets * 2 : tickets;

            if (!combinedOrdersMap.has(id)) {
              combinedOrdersMap.set(id, { id, rev, type, tickets, expectedCount });
              ourTotalRevenue += rev;
            }
          });
        };

        processRows(reports.orders, ordersMap);
        processRows(reports.webhook, webhookMap);

        const afcFreqMap = new Map();
        let afcTotalRevenue = 0;
        const afcRows = reports.afc.slice(afcMap.headerIndex + 1);

        afcRows.forEach(row => {
          const id = row[afcMap.id]?.toString().trim();
          if (!id) return;

          const isDTVM = id.toUpperCase().startsWith('DTVM');
          if (afcFilter === 'DTVM' && !isDTVM) return;
          if (afcFilter === 'NOT_DTVM' && isDTVM) return;

          const rev = parseCurrency(row[afcMap.revenue]);
          afcTotalRevenue += rev;

          const current = afcFreqMap.get(id) || { count: 0, rows: [] };
          current.count += 1;
          current.rows.push(row);
          afcFreqMap.set(id, current);
        });

        const missingInAfc = [];
        const countMismatch = [];

        combinedOrdersMap.forEach((val, id) => {
          const afcData = afcFreqMap.get(id);
          if (!afcData) {
            missingInAfc.push({ 
              id, 
              amount: val.rev, 
              type: val.type, 
              tickets: val.tickets, 
              expected: val.expectedCount, 
              actual: 0,
              reason: 'Missing in AFC'
            });
          } else if (afcData.count !== val.expectedCount) {
            countMismatch.push({
              id,
              amount: val.rev,
              type: val.type,
              tickets: val.tickets,
              expected: val.expectedCount,
              actual: afcData.count,
              reason: `Count mismatch (Expected ${val.expectedCount}, Found ${afcData.count})`
            });
          }
        });

        const missingInOur = [];
        afcFreqMap.forEach((val, id) => {
          if (!combinedOrdersMap.has(id)) {
            missingInOur.push({ id, amount: 0, count: val.count }); 
          }
        });

        setResults({
          missingInAfc: [...missingInAfc, ...countMismatch],
          missingInOur,
          ourTotalRevenue,
          afcTotalRevenue,
          revenueMismatch: ourTotalRevenue - afcTotalRevenue,
          totalOrdersOur: combinedOrdersMap.size,
          totalOrdersAfc: afcFreqMap.size
        });
      } catch (err) {
        console.error("Processing error:", err);
        alert("An error occurred during comparison.");
      } finally {
        setIsProcessing(false);
      }
    }, 100);
  };

  const exportMissing = (data, filename) => {
    if (!data.length) return;
    // Map data to a cleaner format for Excel if needed
    const exportData = data.map(item => ({
      'Order ID': item.id,
      'Ticket Type': item.type,
      'Num of Tickets': item.tickets,
      'Expected in AFC': item.expected,
      'Actual in AFC': item.actual,
      'Amount': item.amount,
      'Reason/Status': item.reason
    }));
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Comparison Results");
    XLSX.writeFile(wb, `${filename}.xlsx`);
  };

  const [searchId, setSearchId] = useState('');
  const [searchResult, setSearchResult] = useState(null);

  const handleSearch = () => {
    if (!reports.afc) {
      alert("Please upload the AFC Report first.");
      return;
    }
    const afcKeywords = {
      id: ['merchant_order_id', 'merchant_order', 'merchant order', 'order_id', 'orderid'],
      revenue: ['actual_fare', 'actual fare', 'fare', 'amount'],
      type: ['ticket_type', 'ticket type', 'type'],
      status: ['ticket_status', 'ticket status', 'status']
    };

    const afcMap = detectHeaders(reports.afc, afcKeywords);

    if (afcMap.headerIndex === -1) {
      alert("Could not detect headers in AFC file.");
      return;
    }

    const afcRows = reports.afc.slice(afcMap.headerIndex + 1);
    const target = searchId.trim().toLowerCase();
    
    const matches = afcRows.filter(row => {
      const id = row[afcMap.id]?.toString().trim().toUpperCase();
      const isDTVM = id.startsWith('DTVM');
      const matchesFilter = afcFilter === 'DTVM' ? isDTVM : !isDTVM;
      return id === target.toUpperCase() && matchesFilter;
    });

    if (matches.length > 0) {
      const first = matches[0];
      setSearchResult({
        found: true,
        count: matches.length,
        id: first[afcMap.id],
        amount: parseCurrency(first[afcMap.revenue]),
        type: first[afcMap.type] || 'N/A',
        status: first[afcMap.status] || 'N/A'
      });
    } else {
      setSearchResult({ found: false, id: searchId });
    }
  };

  return (
    <div className="space-y-12 animate-in fade-in slide-in-from-right-4 duration-500">
      <header className="space-y-6">
        <Link to="/" className="inline-flex items-center gap-2 text-slate-400 hover:text-white transition-colors">
          <ArrowLeft className="w-4 h-4" /> Back to Suite
        </Link>
        <div className="text-center space-y-2">
          <h1 className="text-5xl font-bold bg-gradient-to-br from-white to-primary bg-clip-text text-transparent">
            Compare with AFC
          </h1>
          <p className="text-slate-400 text-lg">Quickly verify Order IDs and reconcile revenue against AFC reports.</p>
        </div>
      </header>

      {/* Manual Search Section */}
      <section className="glass p-8 rounded-3xl space-y-6">
        <div className="flex flex-col md:flex-row gap-4 items-end">
          <div className="flex-1 space-y-2">
            <label className="text-sm font-bold text-slate-400 ml-1">Paste Order ID to Check in AFC</label>
            <input 
              type="text" 
              value={searchId}
              onChange={(e) => setSearchId(e.target.value)}
              placeholder="Enter pg_order_id (e.g. 123456789)"
              className="w-full bg-white/5 border border-white/10 rounded-2xl px-6 py-4 focus:border-primary/50 transition-colors outline-none"
            />
          </div>
          <button 
            onClick={handleSearch}
            className="btn btn-primary px-8 py-4 h-[58px]"
          >
            Check in AFC
          </button>
        </div>

        {searchResult && (
          <div className={`p-6 rounded-2xl animate-in zoom-in-95 duration-300 ${searchResult.found ? 'bg-green-500/10 border border-green-500/20' : 'bg-red-500/10 border border-red-500/20'}`}>
            {searchResult.found ? (
              <div className="flex items-center gap-4">
                <CheckCircle2 className="w-8 h-8 text-green-400" />
                <div>
                  <h3 className="font-bold text-green-400 text-lg">Order Found!</h3>
                  <p className="text-sm text-slate-400">
                    ID: <span className="text-white font-mono">{searchResult.id}</span> | 
                    Found: <span className="text-white font-bold">{searchResult.count} times</span> | 
                    Fare: <span className="text-white font-bold">₹{searchResult.amount}</span> | 
                    Type: <span className="text-white capitalize">{searchResult.type}</span> |
                    Status: <span className={`font-bold ${searchResult.status.toLowerCase().includes('success') || searchResult.status.toLowerCase().includes('paid') ? 'text-green-400' : 'text-yellow-400'}`}>{searchResult.status}</span>
                  </p>
                </div>
              </div>
            ) : (
              <div className="flex items-center gap-4">
                <AlertCircle className="w-8 h-8 text-red-400" />
                <div>
                  <h3 className="font-bold text-red-400 text-lg">Not Found</h3>
                  <p className="text-sm text-slate-400">The Order ID <span className="text-white font-mono">{searchResult.id}</span> does not exist in the uploaded AFC report.</p>
                </div>
              </div>
            )}
          </div>
        )}
      </section>

      <div className="flex items-center gap-4">
        <div className="h-px bg-white/10 flex-1"></div>
        <span className="text-xs font-bold text-slate-500 uppercase tracking-widest">OR Bulk Comparison</span>
        <div className="h-px bg-white/10 flex-1"></div>
      </div>

      {/* Upload Sections */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <div className="relative group">
          <label className={`drop-zone block p-8 ${reports.orders ? 'border-green-500/50 bg-green-500/5' : ''} ${isUploading.orders ? 'opacity-50 pointer-events-none' : ''}`}>
            <input 
              key={reports.orders ? 'orders-loaded' : 'orders-empty'}
              type="file" 
              className="hidden" 
              onChange={(e) => handleFileUpload(e, 'orders')} 
              disabled={isUploading.orders}
            />
            <div className="flex flex-col items-center space-y-3">
              {isUploading.orders ? (
                <div className="w-10 h-10 border-4 border-primary/30 border-t-primary rounded-full animate-spin"></div>
              ) : (
                <FileText className={`w-10 h-10 ${reports.orders ? 'text-green-400' : 'text-primary'}`} />
              )}
              <div className="font-semibold text-lg">{isUploading.orders ? 'Uploading...' : (filenames.orders || 'Orders Report')}</div>
              <p className="text-xs text-slate-500 text-center">pg_order_id, num_of_tickets, ticket_type, total_amount</p>
              {reports.orders && !isUploading.orders && (
                <div className="mt-4 px-3 py-1 rounded-lg text-sm font-medium bg-green-500/20 text-green-400">
                  File Loaded
                </div>
              )}
            </div>
          </label>
          {reports.orders && (
            <button 
              onClick={(e) => { e.preventDefault(); e.stopPropagation(); setReports(prev => ({ ...prev, orders: null })); setFilenames(prev => ({ ...prev, orders: '' })); setResults(null); }}
              className="absolute top-4 right-4 z-20 p-2 bg-red-500/10 hover:bg-red-500/20 text-red-400 rounded-full transition-colors"
              title="Remove File"
            >
              <X className="w-4 h-4" />
            </button>
          )}
        </div>

        <div className="relative group">
          <label className={`drop-zone block p-8 ${reports.webhook ? 'border-green-500/50 bg-green-500/5' : ''} ${isUploading.webhook ? 'opacity-50 pointer-events-none' : ''}`}>
            <input 
              key={reports.webhook ? 'webhook-loaded' : 'webhook-empty'}
              type="file" 
              className="hidden" 
              onChange={(e) => handleFileUpload(e, 'webhook')} 
              disabled={isUploading.webhook}
            />
            <div className="flex flex-col items-center space-y-3">
              {isUploading.webhook ? (
                <div className="w-10 h-10 border-4 border-primary/30 border-t-primary rounded-full animate-spin"></div>
              ) : (
                <FileText className={`w-10 h-10 ${reports.webhook ? 'text-green-400' : 'text-primary'}`} />
              )}
              <div className="font-semibold text-lg">{isUploading.webhook ? 'Uploading...' : (filenames.webhook || 'Webhook Process Report')}</div>
              <p className="text-xs text-slate-500 text-center">pg_order_id, num_of_tickets, ticket_type, total_amount</p>
              {reports.webhook && !isUploading.webhook && (
                <div className="mt-4 px-3 py-1 rounded-lg text-sm font-medium bg-green-500/20 text-green-400">
                  File Loaded
                </div>
              )}
            </div>
          </label>
          {reports.webhook && (
            <button 
              onClick={(e) => { e.preventDefault(); e.stopPropagation(); setReports(prev => ({ ...prev, webhook: null })); setFilenames(prev => ({ ...prev, webhook: '' })); setResults(null); }}
              className="absolute top-4 right-4 z-20 p-2 bg-red-500/10 hover:bg-red-500/20 text-red-400 rounded-full transition-colors"
              title="Remove File"
            >
              <X className="w-4 h-4" />
            </button>
          )}
        </div>

        <div className="relative group">
          <label className={`drop-zone block p-8 ${reports.afc ? 'border-green-500/50 bg-green-500/5' : ''} ${isUploading.afc ? 'opacity-50 pointer-events-none' : ''}`}>
            <input 
              key={reports.afc ? 'afc-loaded' : 'afc-empty'}
              type="file" 
              className="hidden" 
              onChange={(e) => handleFileUpload(e, 'afc')} 
              disabled={isUploading.afc}
            />
            <div className="flex flex-col items-center space-y-3">
              {isUploading.afc ? (
                <div className="w-10 h-10 border-4 border-primary/30 border-t-primary rounded-full animate-spin"></div>
              ) : (
                <Layout className={`w-10 h-10 ${reports.afc ? 'text-green-400' : 'text-primary'}`} />
              )}
              <div className="font-semibold text-lg">{isUploading.afc ? 'Uploading...' : (filenames.afc || 'AFC Report')}</div>
              <p className="text-xs text-slate-500 text-center">MERCHANT_ORDER_ID, ACTUAL_FARE, TICKET_TYPE, TICKET_STATUS</p>
              {reports.afc && !isUploading.afc && (
                <div className="mt-4 px-3 py-1 rounded-lg text-sm font-medium bg-green-500/20 text-green-400">
                  File Loaded
                </div>
              )}
            </div>
          </label>
          {reports.afc && (
            <button 
              onClick={(e) => { e.preventDefault(); e.stopPropagation(); setReports(prev => ({ ...prev, afc: null })); setFilenames(prev => ({ ...prev, afc: '' })); setResults(null); }}
              className="absolute top-4 right-4 z-20 p-2 bg-red-500/10 hover:bg-red-500/20 text-red-400 rounded-full transition-colors"
              title="Remove File"
            >
              <X className="w-4 h-4" />
            </button>
          )}
        </div>
      </div>

      <div className="text-center space-y-6">
        <div className="flex items-center justify-center gap-4">
          <label className="text-sm font-bold text-slate-400 uppercase tracking-widest">AFC Filter:</label>
          <select 
            value={afcFilter} 
            onChange={(e) => setAfcFilter(e.target.value)}
            className="bg-white/5 border border-white/10 rounded-xl px-4 py-2 text-sm focus:border-primary/50 outline-none"
          >
            <option value="DTVM" className="bg-slate-900">Starts with DTVM</option>
            <option value="NOT_DTVM" className="bg-slate-900">Not DTVM</option>
          </select>
        </div>
        
        <button 
          onClick={processComparison}
          disabled={(!reports.orders && !reports.webhook) || !reports.afc || isProcessing}
          className="btn btn-primary px-12 py-4 text-xl flex items-center gap-4 mx-auto"
        >
          {isProcessing && <div className="w-5 h-5 border-2 border-white/30 border-t-white rounded-full animate-spin"></div>}
          {isProcessing ? 'Processing Reports...' : 'Compare Reports'}
        </button>
      </div>

      {results && (
        <section className="space-y-12 animate-in fade-in slide-in-from-bottom-4 duration-500">
          {/* Revenue & Stats Summary */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
            <div className="card">
              <div className="flex items-center gap-3 mb-2">
                <DollarSign className="w-5 h-5 text-primary" />
                <h3 className="text-xs text-slate-500 font-bold tracking-widest uppercase">Our Revenue</h3>
              </div>
              <p className="text-3xl font-bold">₹{results.ourTotalRevenue.toLocaleString()}</p>
              <p className="text-xs text-slate-400 mt-2">{results.totalOrdersOur} Orders</p>
            </div>
            <div className="card">
              <div className="flex items-center gap-3 mb-2">
                <TrendingUp className="w-5 h-5 text-purple-400" />
                <h3 className="text-xs text-slate-500 font-bold tracking-widest uppercase">AFC Revenue</h3>
              </div>
              <p className="text-3xl font-bold">₹{results.afcTotalRevenue.toLocaleString()}</p>
              <p className="text-xs text-slate-400 mt-2">{results.totalOrdersAfc} Orders</p>
            </div>
            <div className="card border-l-4 border-l-red-500">
              <h3 className="text-xs text-slate-500 font-bold tracking-widest uppercase mb-2">Revenue Gap</h3>
              <p className={`text-3xl font-bold ${results.revenueMismatch === 0 ? 'text-green-400' : 'text-red-400'}`}>
                ₹{Math.abs(results.revenueMismatch).toLocaleString()}
              </p>
              <p className="text-xs text-slate-400 mt-2">
                {results.revenueMismatch > 0 ? 'Surplus on our side' : results.revenueMismatch < 0 ? 'Deficit on our side' : 'Perfect Match'}
              </p>
            </div>
            <div className="card">
              <h3 className="text-xs text-slate-500 font-bold tracking-widest uppercase mb-2">Sync Status</h3>
              <div className="flex items-center gap-2">
                {results.missingInAfc.length === 0 && results.missingInOur.length === 0 ? (
                  <>
                    <CheckCircle2 className="w-6 h-6 text-green-400" />
                    <span className="text-green-400 font-bold">In Sync</span>
                  </>
                ) : (
                  <>
                    <AlertCircle className="w-6 h-6 text-red-400" />
                    <span className="text-red-400 font-bold">Mismatch</span>
                  </>
                )}
              </div>
            </div>
          </div>

          {/* Tables Section */}
          <div className="grid grid-cols-1 xl:grid-cols-2 gap-8">
            {/* Missing in AFC */}
            <div className="glass rounded-3xl overflow-hidden">
              <div className="p-6 border-b border-white/10 flex justify-between items-center bg-white/5">
                <div>
                  <h2 className="text-xl font-bold">Missing in AFC</h2>
                  <p className="text-xs text-slate-500">Orders we have, but AFC doesn't.</p>
                </div>
                <button 
                  onClick={() => exportMissing(results.missingInAfc, 'missing_in_afc')}
                  disabled={results.missingInAfc.length === 0}
                  className="btn btn-secondary py-2 px-4 text-xs flex items-center gap-2"
                >
                  <Download className="w-3 h-3" /> Export
                </button>
              </div>
              <div className="overflow-x-auto max-h-[400px]">
                <table className="w-full text-left">
                  <thead className="sticky-header">
                    <tr>
                      <th className="p-4 text-[10px] font-bold text-slate-500 uppercase">Order ID</th>
                      <th className="p-4 text-[10px] font-bold text-slate-500 uppercase">Type</th>
                      <th className="p-4 text-[10px] font-bold text-slate-500 uppercase text-center">Tickets</th>
                      <th className="p-4 text-[10px] font-bold text-slate-500 uppercase text-center">Exp/Act</th>
                      <th className="p-4 text-[10px] font-bold text-slate-500 uppercase">Reason</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-white/5">
                    {results.missingInAfc.length > 0 ? (
                      results.missingInAfc.map((res, i) => (
                        <tr key={i} className="hover:bg-white/5 transition-colors">
                          <td className="p-4 font-mono text-sm">{res.id}</td>
                          <td className="p-4 capitalize text-sm">{res.type}</td>
                          <td className="p-4 text-sm text-center font-bold text-primary">{res.tickets}</td>
                          <td className="p-4 text-sm text-center">
                            <span className="text-slate-400">{res.expected}</span>
                            <span className="mx-1 text-slate-600">/</span>
                            <span className={res.actual === 0 ? 'text-red-400' : 'text-yellow-400'}>{res.actual}</span>
                          </td>
                          <td className="p-4 text-xs font-medium text-red-400/80">{res.reason}</td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan="5" className="p-12 text-center text-slate-500 italic">No missing orders or count mismatches found.</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Missing in Our Report */}
            <div className="glass rounded-3xl overflow-hidden">
              <div className="p-6 border-b border-white/10 flex justify-between items-center bg-white/5">
                <div>
                  <h2 className="text-xl font-bold">Missing in Our Report</h2>
                  <p className="text-xs text-slate-500">Orders AFC has, but we don't.</p>
                </div>
                <button 
                  onClick={() => exportMissing(results.missingInOur, 'missing_in_our_report')}
                  disabled={results.missingInOur.length === 0}
                  className="btn btn-secondary py-2 px-4 text-xs flex items-center gap-2"
                >
                  <Download className="w-3 h-3" /> Export
                </button>
              </div>
              <div className="overflow-x-auto max-h-[400px]">
                <table className="w-full text-left">
                  <thead className="sticky-header">
                    <tr>
                      <th className="p-4 text-[10px] font-bold text-slate-500 uppercase">Order ID</th>
                      <th className="p-4 text-[10px] font-bold text-slate-500 uppercase text-center">Count in AFC</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-white/5">
                    {results.missingInOur.length > 0 ? (
                      results.missingInOur.map((res, i) => (
                        <tr key={i} className="hover:bg-white/5 transition-colors">
                          <td className="p-4 font-mono text-sm">{res.id}</td>
                          <td className="p-4 text-sm text-center font-bold text-purple-400">{res.count}</td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan="2" className="p-12 text-center text-slate-500 italic">No extra orders found in AFC.</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </section>
      )}
    </div>
  );
};

export default AfcComparer;
