/**
 * WorkPilot logo — SVG recreation of the brand asset: compass ring with four
 * star points, gradient paper-plane with speed lines, "Work Pilot" wordmark.
 */

export function LogoMark({ size = 36 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 100 100" fill="none">
      <defs>
        <linearGradient id="wp-blue" x1="0" y1="100" x2="100" y2="0">
          <stop offset="0%" stopColor="#1e3a8a" />
          <stop offset="55%" stopColor="#2563eb" />
          <stop offset="100%" stopColor="#3b82f6" />
        </linearGradient>
        <linearGradient id="wp-plane" x1="20" y1="80" x2="80" y2="20">
          <stop offset="0%" stopColor="#1d4ed8" />
          <stop offset="100%" stopColor="#60a5fa" />
        </linearGradient>
      </defs>
      {/* Compass ring */}
      <circle cx="50" cy="50" r="34" stroke="url(#wp-blue)" strokeWidth="7" fill="none" />
      {/* Star points N E S W */}
      <path d="M50 4 L56 16 L50 13 L44 16 Z" fill="url(#wp-blue)" />
      <path d="M96 50 L84 56 L87 50 L84 44 Z" fill="url(#wp-blue)" />
      <path d="M50 96 L44 84 L50 87 L56 84 Z" fill="url(#wp-blue)" />
      <path d="M4 50 L16 44 L13 50 L16 56 Z" fill="url(#wp-blue)" />
      {/* Paper plane */}
      <path d="M30 70 L78 24 L58 74 L48 56 Z" fill="url(#wp-plane)" />
      <path d="M48 56 L78 24 L52 50 Z" fill="#1e40af" opacity="0.85" />
      {/* Speed lines */}
      <rect x="14" y="66" width="16" height="4.5" rx="2.25" transform="rotate(-32 14 66)" fill="#3b82f6" />
      <rect x="16" y="76" width="12" height="4.5" rx="2.25" transform="rotate(-32 16 76)" fill="#2563eb" />
      <rect x="20" y="85" width="8" height="4.5" rx="2.25" transform="rotate(-32 20 85)" fill="#1d4ed8" />
    </svg>
  );
}

export function LogoFull({ height = 34, light = false }: { height?: number; light?: boolean }) {
  return (
    <span className="inline-flex items-center gap-2.5">
      <LogoMark size={height + 8} />
      <span className="font-bold leading-none tracking-tight" style={{ fontSize: height * 0.62 }}>
        <span className={light ? 'text-white' : 'text-slate-900 dark:text-white'}>Work</span>
        <span className="bg-gradient-to-r from-blue-600 to-blue-400 bg-clip-text text-transparent"> Pilot</span>
      </span>
    </span>
  );
}
