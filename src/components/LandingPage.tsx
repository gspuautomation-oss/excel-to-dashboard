import React from 'react';
import { useNavigate } from 'react-router-dom';
import { ArrowRight, Lock, ShieldCheck, Zap } from 'lucide-react';
import { motion } from 'motion/react';

export const LandingPage: React.FC = () => {
  const navigate = useNavigate();

  return (
    <div className="relative min-h-screen w-full overflow-hidden bg-white flex flex-col justify-between select-none">
      {/* Soft Icy-Blue Ambient Edge Lighting (Top-Left, Top-Right & Lower Glow) */}
      <div
        aria-hidden="true"
        className="pointer-events-none absolute inset-0"
        style={{
          background: `
            radial-gradient(circle at 50% 36%, rgba(255, 255, 255, 1) 0%, rgba(255, 255, 255, 0.96) 42%, rgba(215, 238, 255, 0.52) 76%, rgba(164, 216, 255, 0.72) 100%),
            radial-gradient(ellipse at 12% 88%, rgba(8, 103, 245, 0.28) 0%, transparent 55%),
            radial-gradient(ellipse at 88% 88%, rgba(85, 200, 245, 0.34) 0%, transparent 55%)
          `,
        }}
      />

      {/* Subtle Upper Curved Atmospheric Arcs (Left & Right) */}
      <svg
        aria-hidden="true"
        viewBox="0 0 1600 900"
        preserveAspectRatio="none"
        className="pointer-events-none absolute inset-0 w-full h-full"
      >
        <defs>
          <linearGradient id="upperArcLeft" x1="0%" y1="15%" x2="45%" y2="85%">
            <stop offset="0%" stopColor="#FFFFFF" stopOpacity="0.9" />
            <stop offset="50%" stopColor="#B9E2FF" stopOpacity="0.65" />
            <stop offset="100%" stopColor="#55C8F5" stopOpacity="0.1" />
          </linearGradient>
          <linearGradient id="upperArcRight" x1="100%" y1="15%" x2="55%" y2="85%">
            <stop offset="0%" stopColor="#FFFFFF" stopOpacity="0.9" />
            <stop offset="50%" stopColor="#B9E2FF" stopOpacity="0.65" />
            <stop offset="100%" stopColor="#55C8F5" stopOpacity="0.1" />
          </linearGradient>
        </defs>
        {/* Left Delicate Sweeping Curve */}
        <path
          d="M-60,180 C150,280 280,480 460,720 L-60,720 Z"
          fill="url(#upperArcLeft)"
          opacity="0.45"
        />
        <path
          d="M-40,175 C165,285 295,485 485,735"
          fill="none"
          stroke="#FFFFFF"
          strokeWidth="2.5"
          opacity="0.85"
        />
        {/* Right Delicate Sweeping Curve */}
        <path
          d="M1660,180 C1450,280 1320,480 1140,720 L1660,720 Z"
          fill="url(#upperArcRight)"
          opacity="0.45"
        />
        <path
          d="M1640,175 C1435,285 1305,485 1115,735"
          fill="none"
          stroke="#FFFFFF"
          strokeWidth="2.5"
          opacity="0.85"
        />
      </svg>

      {/* Top Spacer to keep center bright & balanced like Reference Image 1 */}
      <div className="h-12 sm:h-16" />

      {/* CENTRAL HERO CONTENT */}
      <main className="relative z-20 flex-1 flex flex-col items-center justify-center px-6 text-center -mt-6 sm:-mt-10">
        <motion.div
          initial={{ opacity: 0, y: 18 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.75, ease: [0.16, 1, 0.3, 1] }}
          className="max-w-5xl mx-auto flex flex-col items-center"
        >
          {/* Main Luxury Serif Heading */}
          <h1 className="font-serif-luxury tracking-tight leading-[1.06]">
            <span className="block text-6xl sm:text-7xl md:text-[92px] lg:text-[108px] font-bold text-[#003087] drop-shadow-[0_2px_10px_rgba(0,48,135,0.06)]">
              GSPU
            </span>
            <span className="block mt-1 sm:mt-2 text-4xl sm:text-6xl md:text-[74px] lg:text-[86px] font-bold bg-gradient-to-r from-[#0041A8] via-[#0867F5] to-[#0066FF] bg-clip-text text-transparent pb-2">
              Financial Analytics
            </span>
          </h1>

          {/* GET STARTED Pill Button */}
          <motion.div
            initial={{ opacity: 0, scale: 0.96 }}
            animate={{ opacity: 1, scale: 1 }}
            transition={{ duration: 0.65, delay: 0.18, ease: [0.16, 1, 0.3, 1] }}
            className="mt-10 sm:mt-14"
          >
            <button
              type="button"
              onClick={() => navigate('/upload')}
              className="group relative inline-flex items-center justify-between gap-6 pl-10 pr-3.5 py-3.5 sm:pl-12 sm:pr-4 sm:py-4 rounded-full text-white animate-gspu-btn-glow transition-all duration-300 hover:scale-[1.03] active:scale-[0.99] cursor-pointer border border-[#55C8F5]/50"
              style={{
                background:
                  'linear-gradient(180deg, #0747A6 0%, #002F87 48%, #002266 100%)',
              }}
            >
              {/* Subtle top glossy rim */}
              <span
                aria-hidden="true"
                className="pointer-events-none absolute inset-x-5 top-0.5 h-[1.5px] rounded-full bg-gradient-to-r from-transparent via-white/55 to-transparent"
              />
              <span className="font-serif-luxury text-2xl sm:text-[28px] font-semibold tracking-wide text-white pl-1">
                Get Started
              </span>
              <span className="w-10 h-10 sm:w-11 sm:h-11 rounded-full bg-white text-[#003087] flex items-center justify-center shadow-md transition-transform duration-300 group-hover:translate-x-1">
                <ArrowRight className="w-5 h-5 sm:w-6 sm:h-6 stroke-[2.5]" />
              </span>
            </button>
          </motion.div>

          {/* Informational Feature Highlights Under the Button */}
          <motion.div
            initial={{ opacity: 0, y: 12 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.65, delay: 0.32, ease: [0.16, 1, 0.3, 1] }}
            className="mt-9 sm:mt-11 flex flex-wrap items-center justify-center gap-4 sm:gap-7 text-[#003087]"
          >
            {/* Highlight 1: Fast & Accurate */}
            <div className="flex items-center gap-3">
              <span className="w-9 h-9 sm:w-10 sm:h-10 rounded-full bg-white shadow-[0_4px_14px_rgba(8,103,245,0.16)] border border-[#B9E2FF]/80 flex items-center justify-center text-[#0867F5]">
                <Zap className="w-4 h-4 sm:w-5 sm:h-5 fill-[#0867F5]" />
              </span>
              <span className="font-serif-luxury text-base sm:text-xl font-semibold tracking-tight text-[#003087]">
                Fast &amp; Accurate
              </span>
            </div>

            <span
              aria-hidden="true"
              className="hidden sm:inline-block h-6 w-[1px] bg-[#003087]/20"
            />

            {/* Highlight 2: 100% Browser Based */}
            <div className="flex items-center gap-3">
              <span className="w-9 h-9 sm:w-10 sm:h-10 rounded-full bg-white shadow-[0_4px_14px_rgba(8,103,245,0.16)] border border-[#B9E2FF]/80 flex items-center justify-center text-[#0867F5]">
                <ShieldCheck className="w-4 h-4 sm:w-5 sm:h-5 fill-[#0867F5] text-white" />
              </span>
              <span className="font-serif-luxury text-base sm:text-xl font-semibold tracking-tight text-[#003087]">
                100% Browser Based
              </span>
            </div>

            <span
              aria-hidden="true"
              className="hidden sm:inline-block h-6 w-[1px] bg-[#003087]/20"
            />

            {/* Highlight 3: Your Data Stays Local */}
            <div className="flex items-center gap-3">
              <span className="w-9 h-9 sm:w-10 sm:h-10 rounded-full bg-white shadow-[0_4px_14px_rgba(8,103,245,0.16)] border border-[#B9E2FF]/80 flex items-center justify-center text-[#0867F5]">
                <Lock className="w-4 h-4 sm:w-5 sm:h-5 fill-[#0867F5]" />
              </span>
              <span className="font-serif-luxury text-base sm:text-xl font-semibold tracking-tight text-[#003087]">
                Your Data Stays Local
              </span>
            </div>
          </motion.div>
        </motion.div>
      </main>

      {/* LAYERED FLOWING WAVE RIBBONS (Lower-Left & Lower-Right Sweeping to Center) */}
      <div
        aria-hidden="true"
        className="pointer-events-none relative w-full h-44 sm:h-56 md:h-64 -mt-20 sm:-mt-28 z-10 overflow-hidden"
      >
        {/* Ambient Cyan Center-Bottom Horizon Glow */}
        <div
          className="absolute bottom-2 left-1/2 -translate-x-1/2 w-[520px] h-20 rounded-full blur-2xl"
          style={{
            background:
              'radial-gradient(circle, rgba(255,255,255,0.95) 0%, rgba(85,200,245,0.45) 55%, transparent 100%)',
          }}
        />

        {/* Wave Layer 1 (Slow Flowing Deep & Royal Blue Ribbons) */}
        <svg
          viewBox="0 0 1600 360"
          preserveAspectRatio="none"
          className="absolute inset-0 w-full h-full animate-gspu-wave-slow"
        >
          <defs>
            <linearGradient id="waveLeftRoyal" x1="0%" y1="10%" x2="75%" y2="95%">
              <stop offset="0%" stopColor="#0867F5" stopOpacity="0.92" />
              <stop offset="45%" stopColor="#55C8F5" stopOpacity="0.65" />
              <stop offset="100%" stopColor="#B9E2FF" stopOpacity="0.15" />
            </linearGradient>
            <linearGradient id="waveRightRoyal" x1="100%" y1="10%" x2="25%" y2="95%">
              <stop offset="0%" stopColor="#0867F5" stopOpacity="0.9" />
              <stop offset="45%" stopColor="#55C8F5" stopOpacity="0.68" />
              <stop offset="100%" stopColor="#B9E2FF" stopOpacity="0.15" />
            </linearGradient>
            <linearGradient id="waveBaseHorizon" x1="0%" y1="0%" x2="100%" y2="0%">
              <stop offset="0%" stopColor="#55C8F5" stopOpacity="0.55" />
              <stop offset="50%" stopColor="#E6F4FF" stopOpacity="0.85" />
              <stop offset="100%" stopColor="#55C8F5" stopOpacity="0.55" />
            </linearGradient>
          </defs>

          {/* Left Sweeping Ribbon */}
          <path
            d="M-80,40 C120,145 310,245 780,315 L780,360 L-80,360 Z"
            fill="url(#waveLeftRoyal)"
          />
          <path
            d="M-60,125 C150,215 390,285 880,330 L880,360 L-60,360 Z"
            fill="#0867F5"
            opacity="0.35"
          />

          {/* Right Sweeping Ribbon */}
          <path
            d="M1680,40 C1480,145 1290,245 820,315 L820,360 L1680,360 Z"
            fill="url(#waveRightRoyal)"
          />
          <path
            d="M1660,125 C1450,215 1210,285 720,330 L720,360 L1660,360 Z"
            fill="#0867F5"
            opacity="0.35"
          />

          {/* Soft Horizon Base */}
          <path
            d="M0,295 Q800,255 1600,295 L1600,360 L0,360 Z"
            fill="url(#waveBaseHorizon)"
          />
        </svg>

        {/* Wave Layer 2 (Counter-Flowing Cyan & Luminous White Highlight Curves) */}
        <svg
          viewBox="0 0 1600 360"
          preserveAspectRatio="none"
          className="absolute inset-0 w-full h-full animate-gspu-wave-reverse"
        >
          <defs>
            <linearGradient id="cyanRibbonLeft" x1="0%" y1="30%" x2="65%" y2="90%">
              <stop offset="0%" stopColor="#0052CC" stopOpacity="0.85" />
              <stop offset="55%" stopColor="#55C8F5" stopOpacity="0.55" />
              <stop offset="100%" stopColor="#FFFFFF" stopOpacity="0.0" />
            </linearGradient>
            <linearGradient id="cyanRibbonRight" x1="100%" y1="30%" x2="35%" y2="90%">
              <stop offset="0%" stopColor="#0052CC" stopOpacity="0.85" />
              <stop offset="55%" stopColor="#55C8F5" stopOpacity="0.55" />
              <stop offset="100%" stopColor="#FFFFFF" stopOpacity="0.0" />
            </linearGradient>
          </defs>

          {/* Left Dynamic Ribbon Fold */}
          <path
            d="M-40,145 C140,210 295,275 690,320 C410,305 180,260 -40,215 Z"
            fill="url(#cyanRibbonLeft)"
          />
          {/* Right Dynamic Ribbon Fold */}
          <path
            d="M1640,145 C1460,210 1305,275 910,320 C1190,305 1420,260 1640,215 Z"
            fill="url(#cyanRibbonRight)"
          />

          {/* Luminous White Crest Lines */}
          <path
            d="M-30,45 C160,155 360,248 800,312 C1240,248 1440,155 1630,45"
            fill="none"
            stroke="#FFFFFF"
            strokeWidth="2.5"
            opacity="0.9"
          />
          <path
            d="M-30,135 C190,230 430,290 800,325 C1170,290 1410,230 1630,135"
            fill="none"
            stroke="#FFFFFF"
            strokeWidth="2"
            opacity="0.85"
          />
          <path
            d="M-20,245 C290,295 540,318 800,322 C1060,318 1310,295 1620,245"
            fill="none"
            stroke="#FFFFFF"
            strokeWidth="3"
            opacity="0.95"
          />
        </svg>
      </div>
    </div>
  );
};
