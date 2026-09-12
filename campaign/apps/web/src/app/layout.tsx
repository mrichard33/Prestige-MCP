import type { Metadata, Viewport } from 'next';
import Link from 'next/link';
import './globals.css';
import { getSession } from '@/lib/auth';
import { asService } from '@/lib/db';
import { signOutAction } from '@/lib/actions/system';

export const metadata: Metadata = {
  title: 'Campaign Console',
  description: 'Approve, schedule and monitor email campaigns sent through Microsoft Graph.',
};

// Phones are a first-class way to check on a running campaign, so the console
// is laid out to fit one without anything being clipped off the side.
export const viewport: Viewport = {
  width: 'device-width',
  initialScale: 1,
};

export const dynamic = 'force-dynamic';

const NAV = [
  { href: '/', label: 'Dashboard' },
  { href: '/campaigns', label: 'Campaigns' },
  { href: '/contacts', label: 'Contacts' },
  { href: '/suppressions', label: 'Suppressions' },
  { href: '/audit', label: 'Audit' },
  { href: '/settings', label: 'Settings' },
];

async function globalState() {
  try {
    return await asService(async (client) => {
      const { rows } = await client.query<{
        emergency_stop: boolean;
        global_send_enabled: boolean;
        production_mode: boolean;
        emergency_stop_reason: string | null;
      }>(
        `SELECT emergency_stop, global_send_enabled, production_mode, emergency_stop_reason
           FROM campaign.system_controls WHERE id`,
      );
      return rows[0] ?? null;
    });
  } catch {
    return null;
  }
}

export default async function RootLayout({ children }: { children: React.ReactNode }) {
  const session = await getSession();
  const controls = await globalState();

  return (
    <html lang="en">
      <body className="min-h-screen">
        {/* The emergency stop is the first thing on the page, always. */}
        {controls?.emergency_stop && (
          <div className="stop-banner px-4 py-2 text-center text-sm">
            EMERGENCY STOP ENGAGED — nothing will send.
            {controls.emergency_stop_reason ? ` Reason: ${controls.emergency_stop_reason}` : ''}
          </div>
        )}
        {controls && !controls.emergency_stop && !controls.global_send_enabled && (
          <div className="bg-amber-600 px-4 py-2 text-center text-sm font-semibold text-white">
            Global sending is switched off. Campaigns will not send.
          </div>
        )}

        {session && (
          <header className="border-b border-slate-200 bg-white">
            {/* One wrapping row. On a phone the navigation drops to its own
                full-width line and wraps rather than scrolling, so every
                section stays visible instead of running off the edge. */}
            <div className="mx-auto flex max-w-7xl flex-wrap items-center gap-x-3 gap-y-2 px-4 py-3 md:gap-x-4">
              <span className="shrink-0 text-sm font-semibold text-slate-900 md:text-base">
                Campaign Console
              </span>
              <nav className="order-last w-full md:order-none md:w-auto md:flex-1">
                <div className="flex flex-wrap gap-1">
                  {NAV.map((item) => (
                    <Link
                      key={item.href}
                      href={item.href}
                      className="shrink-0 whitespace-nowrap rounded px-2.5 py-1.5 text-sm text-slate-600 hover:bg-slate-100 hover:text-slate-900 md:px-3"
                    >
                      {item.label}
                    </Link>
                  ))}
                </div>
              </nav>
              <span
                className={`ml-auto shrink-0 rounded border px-2 py-0.5 text-xs font-medium md:ml-0 ${
                  controls?.production_mode
                    ? 'border-red-200 bg-red-50 text-red-800'
                    : 'border-slate-200 bg-slate-50 text-slate-600'
                }`}
                title={
                  controls?.production_mode
                    ? 'Production sending is enabled. Campaigns in production mode can reach real recipients.'
                    : 'Production mode is off. Only addresses on the test allowlist can be reached.'
                }
              >
                {controls?.production_mode ? 'PRODUCTION' : 'TEST MODE'}
              </span>
              {/* Identity is the first thing to give up room for: below md it
                  would push the sign-out control off the edge. */}
              <span
                className="hidden max-w-[16rem] truncate text-xs text-slate-500 md:inline"
                title={`${session.email} · ${session.role}`}
              >
                {session.email} · {session.role}
              </span>
              <form action={signOutAction} className="shrink-0">
                <button className="text-xs text-slate-500 underline hover:text-slate-800">
                  Sign out
                </button>
              </form>
            </div>
          </header>
        )}

        <main className="mx-auto max-w-7xl px-4 py-6">{children}</main>
      </body>
    </html>
  );
}
