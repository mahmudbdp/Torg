import React, { useEffect } from 'react';
import { Link } from 'react-router-dom';
import { DotLottieReact } from '@lottiefiles/dotlottie-react';
import { FileSpreadsheet, Clock, ArrowRight } from 'lucide-react';

const Home = () => {
  useEffect(() => {
    document.title = "Home | TORG ODP JOB";
  }, []);

  return (
    <div className="animate-fade-in">
      {/* Hero Section */}
      <section className="relative py-24 px-6 overflow-hidden border-b border-gray-200 bg-gradient-to-br from-white via-blue-50 to-orange-50/30">
        <div className="max-w-6xl mx-auto flex flex-col md:flex-row items-center justify-between gap-12 relative z-10">
          <div className="flex-1 text-aws-navy max-w-2xl">
            <h2 className="text-5xl font-extrabold mb-6 leading-tight tracking-tight">
              Everything you need, <br />
              <span className="text-aws-orange">running locally</span>
            </h2>
            <p className="text-lg text-gray-600 mb-10 leading-relaxed max-w-lg">
              These tools process your data securely inside your browser. No files are uploaded to any external servers. Experience lightning fast, private utility apps.
            </p>
            <div className="flex gap-4">
              <Link 
                to="/converter" 
                className="bg-aws-orange hover:bg-orange-600 text-white font-bold py-3 px-8 rounded-md transition-all shadow-lg hover:shadow-orange-500/30 flex items-center gap-2"
              >
                Launch Converter <ArrowRight size={20} />
              </Link>
            </div>
          </div>
          <div className="flex-shrink-0 w-80 h-80 md:w-96 md:h-96 relative">
            <div className="absolute inset-0 bg-blue-500/5 blur-3xl rounded-full"></div>
            <DotLottieReact
              src="https://lottie.host/95198946-ca04-452d-b3f2-4805ec36b0cb/OP3toB0OEo.lottie"
              loop
              autoplay
              className="w-full h-full relative z-10 drop-shadow-xl"
            />
          </div>
        </div>
      </section>

      {/* Tools Grid */}
      <section className="py-20 px-6 max-w-6xl mx-auto">
        <h3 className="text-2xl font-bold text-gray-900 mb-8 border-b border-gray-200 pb-4">Available Utilities</h3>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8">
          
          <Link to="/converter" className="group">
            <div className="glass-card h-full p-8 transition-all duration-300 hover:-translate-y-1 hover:border-aws-blue hover:shadow-2xl">
              <div className="h-32 w-full mb-6 flex items-center justify-center rounded-xl bg-blue-50/50 group-hover:bg-blue-50 transition-colors">
                <DotLottieReact
                  src="https://lottie.host/611f6300-561e-4a9c-add6-2299e5eeaa1a/TOE0YmGCKt.lottie"
                  loop
                  autoplay
                  className="w-full h-full transform group-hover:scale-110 transition-transform duration-500"
                />
              </div>
              <h4 className="text-xl font-bold text-gray-900 mb-3">ShipStation to Bluecherry</h4>
              <p className="text-gray-600 mb-6 text-sm leading-relaxed flex-grow">
                Upload a ShipStation order export and convert it into a Bluecherry-ready MJSO .xls file with intelligent SKU parsing and mapping.
              </p>
              <span className="text-aws-blue font-semibold text-sm flex items-center gap-1 group-hover:gap-2 transition-all">
                Open Tool <ArrowRight size={16} />
              </span>
            </div>
          </Link>

          <Link to="/inventory-checker" className="group">
            <div className="glass-card h-full p-8 transition-all duration-300 hover:-translate-y-1 hover:border-aws-orange hover:shadow-2xl">
              <div className="w-14 h-14 bg-orange-50 text-aws-orange rounded-xl flex items-center justify-center mb-6 group-hover:scale-110 transition-transform">
                <Clock size={28} />
              </div>
              <h4 className="text-xl font-bold text-gray-900 mb-3">Inventory Checker</h4>
              <p className="text-gray-600 mb-6 text-sm leading-relaxed flex-grow">
                Merge Open Pick Sales Orders with current Inventory Quantities by Location to verify availability in real-time.
              </p>
              <span className="text-aws-orange font-semibold text-sm flex items-center gap-1 group-hover:gap-2 transition-all">
                Open Tool <ArrowRight size={16} />
              </span>
            </div>
          </Link>

        </div>
      </section>
    </div>
  );
};

export default Home;
