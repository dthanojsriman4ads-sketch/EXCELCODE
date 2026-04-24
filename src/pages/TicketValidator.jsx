import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileText, Layout, Download, ArrowLeft } from 'lucide-react';
import { Link } from 'react-router-dom';

const TicketValidator = () => {
  const [mode, setMode] = useState('single');
  const [results, setResults] = useState([]);
  const [dualData, setDualData] = useState({ orders: null, tickets: null });

  const stats = useMemo(() => {
    const total = results.length;
    const matched = results.filter(r => r.status === 'Match').length;
    const unmatched = total - matched;
    return { total, matched, unmatched };
  }, [results]);

  const handleFileUpload = (e, type) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

      if (mode === 'single') {
        processSingle(data);
      } else {
        setDualData(prev => ({ ...prev, [type]: data }));
      }
    };
    reader.readAsBinaryString(file);
  };

  const processSingle = (data) => {
    const colMap = detectColumns(data, ['pg_order', 'num', 'order_id', 'type']);
    const rows = data.slice(colMap.headerIndex + 1);
    const frequencies = countFrequencies(rows, colMap.order_id);
    runValidation(rows, frequencies, colMap);
  };

  const processDual = () => {
    if (!dualData.orders || !dualData.tickets) return;
    const ordersColMap = detectColumns(dualData.orders, ['pg_order', 'num', 'type']);
    const ordersRows = dualData.orders.slice(ordersColMap.headerIndex + 1);
    const ticketsColMap = detectColumns(dualData.tickets, ['order_id']);
    const ticketsRows = dualData.tickets.slice(ticketsColMap.headerIndex + 1);
    const frequencies = countFrequencies(ticketsRows, ticketsColMap.order_id);
    runValidation(ordersRows, frequencies, ordersColMap);
  };

  const detectColumns = (data, keywords) => {
    let colMap = { pg_order_id: 0, num_of_tickets: 2, order_id: 3, ticket_type: 1, headerIndex: -1 };
    for (let i = 0; i < Math.min(data.length, 10); i++) {
      const row = (data[i] || []).map(c => c?.toString().toLowerCase().trim() || "");
      if (row.some(c => keywords.some(k => c.includes(k)))) {
        colMap.headerIndex = i;
        row.forEach((cell, idx) => {
          if (cell.includes('pg_order')) colMap.pg_order_id = idx;
          else if (cell.includes('num') || (cell.includes('ticket') && !cell.includes('type'))) colMap.num_of_tickets = idx;
          else if (cell.includes('order_id') && !cell.includes('pg')) colMap.order_id = idx;
          else if (cell.includes('type')) colMap.ticket_type = idx;
        });
        break;
      }
    }
    return colMap;
  };

  const countFrequencies = (rows, colIndex) => {
    const freq = {};
    rows.forEach(row => {
      const val = row[colIndex]?.toString().trim();
      if (val) freq[val] = (freq[val] || 0) + 1;
    });
    return freq;
  };

  const runValidation = (targetRows, frequencies, colMap) => {
    const res = [];
    const processed = new Set();
    targetRows.forEach(row => {
      const id = row[colMap.pg_order_id]?.toString().trim();
      const base = parseFloat(row[colMap.num_of_tickets]);
      const type = (row[colMap.ticket_type]?.toString().trim() || 'single').toLowerCase();
      if (!id || isNaN(base)) return;
      if (processed.has(id)) return;
      processed.add(id);
      const actual = frequencies[id] || 0;
      const expected = type.includes('return') ? base * 2 : base;
      const isMatch = Math.abs(actual - expected) < 0.1;
      res.push({ id, type: type.includes('return') ? 'return' : 'single', base, expected, actual, status: isMatch ? 'Match' : 'Mismatch' });
    });
    setResults(res);
  };

  const exportUnmatched = () => {
    const unmatched = results.filter(r => r.status === 'Mismatch');
    if (!unmatched.length) return;
    const ws = XLSX.utils.json_to_sheet(unmatched);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Unmatched");
    XLSX.writeFile(wb, "unmatched_orders.xlsx");
  };

  return (
    <div className="space-y-12 animate-in fade-in slide-in-from-right-4 duration-500">
      <header className="space-y-6">
        <Link to="/" className="inline-flex items-center gap-2 text-slate-400 hover:text-white transition-colors">
          <ArrowLeft className="w-4 h-4" /> Back to Suite
        </Link>
        <div className="text-center space-y-2">
          <h1 className="text-5xl font-bold bg-gradient-to-br from-white to-primary bg-clip-text text-transparent">
            Duplicate Tickets Validator
          </h1>
          <p className="text-slate-400 text-lg">Compare order reports against ticket listings instantly.</p>
        </div>
      </header>

      {/* Mode Switcher */}
      <div className="flex glass p-1.5 rounded-2xl w-fit mx-auto gap-1">
        <button 
          onClick={() => { setMode('single'); setResults([]); }}
          className={`px-6 py-2 rounded-xl font-semibold transition-all ${mode === 'single' ? 'bg-primary text-white shadow-lg shadow-primary/20' : 'text-slate-400 hover:text-white'}`}
        >
          Single File
        </button>
        <button 
          onClick={() => { setMode('dual'); setResults([]); }}
          className={`px-6 py-2 rounded-xl font-semibold transition-all ${mode === 'dual' ? 'bg-primary text-white shadow-lg shadow-primary/20' : 'text-slate-400 hover:text-white'}`}
        >
          Two Reports
        </button>
      </div>

      {/* Upload Sections */}
      <div className="space-y-6">
        {mode === 'single' ? (
          <label className="drop-zone block p-12 group">
            <input type="file" className="hidden" onChange={(e) => handleFileUpload(e, 'single')} />
            <div className="flex flex-col items-center space-y-4">
              <Upload className="w-16 h-16 text-primary group-hover:scale-110 transition-transform" />
              <div className="text-xl">Drag & drop combined file or <span className="text-primary font-bold underline">browse</span></div>
              <p className="text-sm text-slate-500">Requires pg_order_id, num_of_tickets, order_id, ticket_type</p>
            </div>
          </label>
        ) : (
          <div className="space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <label className="drop-zone p-8 group">
                <input type="file" className="hidden" onChange={(e) => handleFileUpload(e, 'orders')} />
                <div className="flex flex-col items-center space-y-3">
                  <FileText className="w-10 h-10 text-primary group-hover:scale-110 transition-transform" />
                  <div className="font-semibold text-lg">Orders Report</div>
                  <p className="text-xs text-slate-500">ID, Count, Type</p>
                  <div className={`mt-4 px-3 py-1 rounded-lg text-sm font-medium ${dualData.orders ? 'bg-green-500/20 text-green-400' : 'bg-white/5 text-slate-500'}`}>
                    {dualData.orders ? 'File Loaded' : 'Pending...'}
                  </div>
                </div>
              </label>
              <label className="drop-zone p-8 group">
                <input type="file" className="hidden" onChange={(e) => handleFileUpload(e, 'tickets')} />
                <div className="flex flex-col items-center space-y-3">
                  <Layout className="w-10 h-10 text-primary group-hover:scale-110 transition-transform" />
                  <div className="font-semibold text-lg">Tickets Report</div>
                  <p className="text-xs text-slate-500">Order ID (Scan)</p>
                  <div className={`mt-4 px-3 py-1 rounded-lg text-sm font-medium ${dualData.tickets ? 'bg-green-500/20 text-green-400' : 'bg-white/5 text-slate-500'}`}>
                    {dualData.tickets ? 'File Loaded' : 'Pending...'}
                  </div>
                </div>
              </label>
            </div>
            <div className="text-center">
              <button 
                onClick={processDual}
                disabled={!dualData.orders || !dualData.tickets}
                className="btn btn-primary px-12 py-4 text-xl"
              >
                Process Reports
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Results Section */}
      {results.length > 0 && (
        <section className="space-y-8">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            <div className="card">
              <h3 className="text-xs text-slate-500 font-bold tracking-widest uppercase mb-1">Total Orders</h3>
              <p className="text-4xl font-bold">{stats.total}</p>
            </div>
            <div className="card">
              <h3 className="text-xs text-slate-500 font-bold tracking-widest uppercase mb-1">Matched</h3>
              <p className="text-4xl font-bold text-green-400">{stats.matched}</p>
            </div>
            <div className="card">
              <h3 className="text-xs text-slate-500 font-bold tracking-widest uppercase mb-1">Unmatched</h3>
              <p className="text-4xl font-bold text-red-400">{stats.unmatched}</p>
            </div>
          </div>

          <div className="glass rounded-3xl overflow-hidden">
            <div className="p-6 border-b border-white/10 flex justify-between items-center">
              <h2 className="text-xl font-bold">Validation Details</h2>
              <button onClick={exportUnmatched} className="btn btn-secondary flex items-center gap-2">
                <Download className="w-4 h-4" /> Export Unmatched
              </button>
            </div>
            <div className="overflow-x-auto max-h-[500px]">
              <table className="w-full text-left">
                <thead className="bg-white/5 sticky top-0 z-10">
                  <tr>
                    <th className="p-4 text-xs font-bold text-slate-500 uppercase">PG Order ID</th>
                    <th className="p-4 text-xs font-bold text-slate-500 uppercase">Type</th>
                    <th className="p-4 text-xs font-bold text-slate-500 uppercase">Expected (Total)</th>
                    <th className="p-4 text-xs font-bold text-slate-500 uppercase">Actual Count</th>
                    <th className="p-4 text-xs font-bold text-slate-500 uppercase">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-white/5">
                  {results.map((res, i) => (
                    <tr key={i} className="hover:bg-white/5 transition-colors">
                      <td className="p-4 font-mono text-sm">{res.id}</td>
                      <td className="p-4 capitalize text-sm">{res.type}</td>
                      <td className="p-4 text-sm">
                        {res.base} <span className="opacity-50 text-xs">(Total: {res.expected})</span>
                      </td>
                      <td className="p-4 text-sm font-bold">{res.actual}</td>
                      <td className="p-4">
                        <span className={`status-badge ${res.status === 'Match' ? 'bg-green-500/10 text-green-400' : 'bg-red-500/10 text-red-400'}`}>
                          {res.status}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </section>
      )}
    </div>
  );
};

export default TicketValidator;
