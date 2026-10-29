import React from 'react';

interface GspuBrandLogoProps {
  compact?: boolean;
  lightText?: boolean;
  iconOnly?: boolean;
  whiteTile?: boolean;
  onClick?: () => void;
}

export const GspuBrandLogo: React.FC<GspuBrandLogoProps> = ({
  compact = false,
  lightText = false,
  iconOnly = false,
  whiteTile = false,
  onClick,
}) => {
  const size = compact ? 38 : 48;

  const emblem = (
    <svg
      width={size}
      height={size}
      viewBox="0 0 56 56"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
      className="shrink-0"
      aria-label="GSPU Financial Analytics Logo"
    >
      <defs>
        <linearGradient id="gspuDiagonalGrad" x1="8" y1="44" x2="46" y2="12" gradientUnits="userSpaceOnUse">
          <stop offset="0%" stopColor="#50AE5D" />
          <stop offset="45%" stopColor="#12A89D" />
          <stop offset="100%" stopColor="#0867F5" />
        </linearGradient>
        <linearGradient id="gspuBarBlue" x1="0" y1="16" x2="0" y2="46" gradientUnits="userSpaceOnUse">
          <stop offset="0%" stopColor="#0867F5" />
          <stop offset="100%" stopColor="#003087" />
        </linearGradient>
        <linearGradient id="gspuTileDark" x1="0" y1="0" x2="56" y2="56" gradientUnits="userSpaceOnUse">
          <stop offset="0%" stopColor="#0B2540" />
          <stop offset="100%" stopColor="#001F54" />
        </linearGradient>
      </defs>

      {whiteTile ? (
        <rect
          x="1"
          y="1"
          width="54"
          height="54"
          rx="14"
          fill="#FFFFFF"
          stroke="#E2EEFF"
          strokeWidth="1.5"
        />
      ) : (
        <rect width="56" height="56" rx="13" fill="url(#gspuTileDark)" />
      )}

      {/* Ascending Diagonal Financial Growth Wing */}
      <path
        d="M12 37L42 13V22L12 43V37Z"
        fill="url(#gspuDiagonalGrad)"
        opacity="0.92"
      />

      {/* Four Ascending Financial Bars */}
      <rect x="13" y="34" width="5.5" height="10" rx="1.5" fill="#50AE5D" />
      <rect x="21" y="29" width="5.5" height="15" rx="1.5" fill="#159E96" />
      <rect x="29" y="23" width="5.5" height="21" rx="1.5" fill="#0867F5" />
      <rect x="37" y="16" width="5.5" height="28" rx="1.5" fill="url(#gspuBarBlue)" />

      {/* Upward Trend Curve & Arrow */}
      <path
        d="M15 39C22 36 30 30 39 20"
        stroke={whiteTile ? '#FFFFFF' : '#55C8F5'}
        strokeWidth="2.6"
        strokeLinecap="round"
      />
      <path
        d="M34.5 19.5H39.5V24.5"
        stroke={whiteTile ? '#FFFFFF' : '#55C8F5'}
        strokeWidth="2.6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );

  if (iconOnly) {
    return (
      <div
        onClick={onClick}
        className={`inline-flex items-center justify-center select-none ${
          onClick ? 'cursor-pointer' : ''
        } ${whiteTile ? 'shadow-[0_6px_20px_rgba(8,103,245,0.10)] rounded-2xl' : ''}`}
      >
        {emblem}
      </div>
    );
  }

  return (
    <div
      onClick={onClick}
      className={`flex items-center gap-3 select-none shrink-0 ${
        onClick ? 'cursor-pointer group' : ''
      }`}
    >
      <div className={whiteTile ? 'shadow-[0_6px_20px_rgba(8,103,245,0.10)] rounded-2xl' : ''}>
        {emblem}
      </div>
      <div className="flex flex-col">
        <span
          className={`font-bold tracking-tight leading-none whitespace-nowrap transition-colors ${
            compact ? 'text-base' : 'text-lg'
          } ${lightText ? 'text-white' : 'text-[#003087]'}`}
        >
          GSPU Financial Analytics
        </span>
        <span
          className={`text-[11px] font-medium tracking-normal mt-0.5 whitespace-nowrap ${
            lightText ? 'text-slate-300' : 'text-slate-500'
          }`}
        >
          TallyPrime Audit · GST · MIS Intelligence
        </span>
      </div>
    </div>
  );
};
