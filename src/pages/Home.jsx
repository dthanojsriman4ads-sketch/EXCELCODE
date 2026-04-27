import React from 'react';
import { Link } from 'react-router-dom';
import { Copy, ChevronRight, Calculator, FileSearch, Zap } from 'lucide-react';

const tools = [
  {
    id: 'ticket-validator',
    title: 'Duplicate Tickets Validator',
    description: 'Verify and cross-match order counts between different Excel reports with custom logic.',
    icon: <Copy className="w-8 h-8 text-primary" />,
    path: '/ticket-validator',
    status: 'Ready'
  },
  {
    id: 'compare-afc',
    title: 'Compare with AFC',
    description: 'Cross-match order IDs and compare revenue between internal reports and AFC reports.',
    icon: <Calculator className="w-8 h-8 text-primary" />,
    path: '/compare-afc',
    status: 'Ready'
  },
  {
    id: 'placeholder-1',
    title: 'Data Formatter',
    description: 'Clean and format messy CSV/Excel data for database imports (Coming Soon).',
    icon: <FileSearch className="w-8 h-8 text-slate-500" />,
    path: '/',
    status: 'Upcoming'
  },
  {
    id: 'placeholder-2',
    title: 'Instant Reporter',
    description: 'Generate PDF reports from raw transaction data in seconds (Coming Soon).',
    icon: <Zap className="w-8 h-8 text-slate-500" />,
    path: '/',
    status: 'Upcoming'
  }
];

const Home = () => {
  return (
    <div className="space-y-12 animate-in fade-in duration-700">
      <header className="text-center space-y-4">
        <div className="flex justify-center">
          <div className="p-3 bg-primary/10 rounded-2xl">
            <Zap className="w-8 h-8 text-primary" />
          </div>
        </div>
        <h1 className="text-6xl font-bold bg-gradient-to-br from-white to-primary bg-clip-text text-transparent">
          Utility Suite
        </h1>
        {/* <p className="text-slate-400 text-xl max-w-2xl mx-auto">
          Powerful, browser-based tools to streamline your administration and data validation workflows.
        </p> */}
      </header>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8">
        {tools.map((tool) => (
          <Link 
            key={tool.id}
            to={tool.path}
            className={`group card text-left space-y-6 transition-all hover:scale-[1.03] hover:border-primary/30 ${tool.status === 'Upcoming' ? 'opacity-70 cursor-not-allowed' : ''}`}
          >
            <div className="flex justify-between items-start">
              <div className="p-4 bg-white/5 rounded-2xl group-hover:bg-primary/10 transition-colors">
                {tool.icon}
              </div>
              <span className={`px-3 py-1 rounded-full text-[10px] font-bold uppercase tracking-widest ${tool.status === 'Ready' ? 'bg-primary/20 text-primary' : 'bg-white/5 text-slate-500'}`}>
                {tool.status}
              </span>
            </div>
            
            <div className="space-y-2">
              <h2 className="text-2xl font-bold group-hover:text-primary transition-colors">{tool.title}</h2>
              <p className="text-slate-400 text-sm leading-relaxed">{tool.description}</p>
            </div>

            <div className="flex items-center text-primary font-bold text-sm gap-1 group-hover:gap-2 transition-all">
              {tool.status === 'Ready' ? 'Launch Tool' : 'Coming Soon'} <ChevronRight className="w-4 h-4" />
            </div>
          </Link>
        ))}
      </div>
    </div>
  );
};

export default Home;
