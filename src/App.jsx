import React, { useEffect } from 'react';
import { HashRouter, Routes, Route, Link, useLocation } from 'react-router-dom';
import Home from './Home';
import Converter from './Converter';
import InventoryChecker from './InventoryChecker';
import { Hexagon } from 'lucide-react';

function ScrollToTop() {
  const { pathname } = useLocation();
  useEffect(() => {
    window.scrollTo(0, 0);
  }, [pathname]);
  return null;
}

function App() {
  return (
    <HashRouter>
      <ScrollToTop />
      <div className="min-h-screen bg-gray-50 flex flex-col">
        {/* Navigation Bar */}
        <nav className="bg-white border-b border-gray-200 sticky top-0 z-50 shadow-sm">
          <div className="max-w-6xl mx-auto px-6 h-16 flex items-center">
            <Link to="/" className="flex items-center gap-3 hover:opacity-80 transition-opacity">
              <div className="w-8 h-8 bg-aws-orange rounded-md flex items-center justify-center text-white shadow-sm">
                <Hexagon size={20} className="fill-white" />
              </div>
              <div>
                <h1 className="text-lg font-bold text-gray-900 leading-tight">Bluecherry Tools</h1>
                <p className="text-[11px] text-gray-500 leading-tight">Local browser-based utilities</p>
              </div>
            </Link>
          </div>
        </nav>

        {/* Main Content Area */}
        <main className="flex-grow">
          <Routes>
            <Route path="/" element={<Home />} />
            <Route path="/converter" element={<Converter />} />
            <Route path="/inventory-checker" element={<InventoryChecker />} />
          </Routes>
        </main>
      </div>
    </HashRouter>
  );
}

export default App;
