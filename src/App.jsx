import React from 'react';
import { BrowserRouter as Router, Routes, Route } from 'react-router-dom';
import Home from './pages/Home';
import TicketValidator from './pages/TicketValidator';
import AfcComparer from './pages/AfcComparer';

const App = () => {
  return (
    <Router>
      <div className="relative min-h-screen flex flex-col items-center py-12 px-4 overflow-x-hidden">
        {/* Global Background Blobs */}
        <div className="fixed inset-0 -z-10 overflow-hidden pointer-events-none">
          <div className="blob bg-primary left-[-100px] top-[-100px] animate-move-slow" />
          <div className="blob bg-pink-500 right-[-100px] bottom-[-100px] animate-move-slower" />
        </div>

        <div className="w-full max-w-6xl space-y-12">
          <Routes>
            <Route path="/" element={<Home />} />
            <Route path="/ticket-validator" element={<TicketValidator />} />
            <Route path="/compare-afc" element={<AfcComparer />} />
          </Routes>
        </div>
        
        <footer className="mt-20 py-8 text-center text-slate-500 text-sm border-t border-white/5 w-full max-w-6xl">
          &copy; {new Date().getFullYear()} Utility Suite. 
        </footer>
      </div>
    </Router>
  );
};

export default App;
