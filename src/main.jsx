import React, { useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import {
  CalendarDays, MapPin, Search, Menu, X, ArrowUpRight, Upload, Download,
  Users, Trophy, Clock3, Info, RotateCcw, Check
} from 'lucide-react';
import './style.css';

/* ------------------------------------------------------------------ *
 *  1. CẤU HÌNH GIẢI — chỉ cần sửa ở đây khi có thông tin chính thức
 * ------------------------------------------------------------------ */
const EVENT = {
  name: 'MSB Pickleball SI Open 2026',
  dateLabel: 'Thứ Bảy',            // ví dụ: 'Thứ Bảy 14.11.2026'
  timeLabel: '07:00 – 11:00',
  venue: 'Địa điểm đang chốt',     // ví dụ: 'Sân Pickleball ...'
  startISO: '',                    // ví dụ: '2026-11-14T07:00:00+07:00' → bật đồng hồ đếm ngược
  prizeTotal: '6.000.000đ'
};

const NAV = [
  ['doi-thi', 'Đội thi đấu'],
  ['lich-trinh', 'Chương trình'],
  ['lich', 'Lịch thi đấu'],
  ['bang', 'Xếp hạng'],
  ['nhanh', 'Nhánh đấu'],
  ['the-le', 'Thể lệ'],
  ['giai-thuong', 'Giải thưởng']
];

const TIMELINE = [
  ['06:30', 'Chuẩn bị', 'Chuẩn bị hậu cần', 'Ban Tổ chức chuẩn bị sân bãi và dụng cụ.'],
  ['07:00', 'Khai mạc', 'Check-in, khai mạc và phổ biến thể lệ', 'Ổn định đội hình, phổ biến quy định thi đấu.'],
  ['07:15', 'Vòng tròn', 'Thi đấu vòng tròn', 'Bốn đội đấu vòng tròn một lượt.'],
  ['09:45', 'Xếp hạng', 'Thi đấu xếp hạng', 'Hạng 1 gặp hạng 2; hạng 3 gặp hạng 4.'],
  ['10:45', 'Trao giải', 'Trao giải và chụp ảnh', 'Dự kiến kết thúc chương trình lúc 11:00.']
];

const PRIZES = [
  { rank: 'Giải Nhất', amount: '3.000.000đ', reward: 'Cúp vô địch', cls: 'first' },
  { rank: 'Giải Nhì', amount: '2.000.000đ', reward: 'Huy chương', cls: 'second' },
  { rank: 'Giải Ba', amount: '1.000.000đ', reward: 'Huy chương', cls: 'third' }
];

const SUB_MATCHES = ['Đôi nam 1', 'Đôi nam 2', 'Đôi nữ', 'Đôi nam nữ'];

const RULES = [
  {
    title: 'Thành phần đội',
    body: <p>Bốn đội thi đấu, mỗi đội có một đội trưởng. Đội hình được cân bằng theo giới tính và trình độ.</p>
  },
  {
    title: 'Thể thức vòng tròn',
    body: (
      <>
        <p>Bốn đội đấu vòng tròn một lượt, không chia bảng. Mỗi cặp đội thi đấu bốn trận nhỏ:</p>
        <ul className="chips-static">{SUB_MATCHES.map(s => <li key={s}>{s}</li>)}</ul>
      </>
    )
  },
  {
    title: 'Tính điểm',
    body: <p>Mỗi trận nhỏ chạm 11 điểm. Thắng hay thua của cặp đội được tính theo số trận nhỏ thắng được.</p>
  },
  {
    title: 'Tranh hạng',
    body: <p>Sau vòng tròn, hạng 1 và hạng 2 tranh Nhất/Nhì; hạng 3 và hạng 4 tranh Ba/Tư.</p>
  }
];

/* ------------------------------------------------------------------ *
 *  2. DỮ LIỆU MẶC ĐỊNH (có thể thay bằng file Excel ở Khu vực BTC)
 * ------------------------------------------------------------------ */
const m = (id, court, time, a, b) => ({
  'Mã trận': id, 'Nội dung': 'Đồng đội', 'Vòng': 'Vòng tròn', 'Bảng': 'Vòng bảng',
  'Sân': court, 'Giờ': time, 'Đội 1': a, 'Đội 2': b, 'Tỉ số': '', 'Trạng thái': 'Dự kiến'
});
const DEFAULT_MATCHES = [
  m('M01', 'Sân 1', '07:30', 'Đội 1 - A Nghi', 'Đội 2 - A Hải'),
  m('M02', 'Sân 2', '07:30', 'Đội 3 - Anh Toàn', 'Đội 4 - Anh Tú'),
  m('M03', 'Sân 1', '08:30', 'Đội 1 - A Nghi', 'Đội 3 - Anh Toàn'),
  m('M04', 'Sân 2', '08:30', 'Đội 2 - A Hải', 'Đội 4 - Anh Tú'),
  m('M05', 'Sân 1', '09:30', 'Đội 1 - A Nghi', 'Đội 4 - Anh Tú'),
  m('M06', 'Sân 2', '09:30', 'Đội 2 - A Hải', 'Đội 3 - Anh Toàn')
];

const roster = (team, captain, ...others) => [
  { 'Đội': team, 'Họ tên': captain, 'Vai trò': 'Đội trưởng' },
  ...others.map(n => ({ 'Đội': team, 'Họ tên': n }))
];
const DEFAULT_PLAYERS = [
  ...roster('Đội 1', 'A Nghi', 'Anh Mạnh Tuấn', 'Anh Việt', 'Anh Thanh Tuấn', 'Lan', 'Vân Anh'),
  ...roster('Đội 2', 'A Hải', 'Trung Anh', 'A Sơn', 'Tú Đinh', 'Chị Chill', 'Chị Phương Bùi'),
  ...roster('Đội 3', 'Anh Toàn', 'Minh Vũ', 'Anh Tiến', 'Chị Phương', 'Chị Nga Phạm'),
  ...roster('Đội 4', 'Anh Tú', 'Anh Nhật', 'A Quang', 'Anh Long')
];

/* ------------------------------------------------------------------ *
 *  3. HÀM TIỆN ÍCH
 * ------------------------------------------------------------------ */
const KEYS = { matches: 'siopen_matches_v2', players: 'siopen_players_v2' };
const readSaved = (key, fallback) => {
  try { const v = localStorage.getItem(key); return v ? JSON.parse(v) : fallback; } catch { return fallback; }
};
const writeSaved = (key, value) => { try { localStorage.setItem(key, JSON.stringify(value)); } catch { /* bỏ qua */ } };
const clearSaved = () => { try { Object.values(KEYS).forEach(k => localStorage.removeItem(k)); } catch { /* bỏ qua */ } };

const byNumeric = (a, b) => String(a).localeCompare(String(b), 'vi', { numeric: true });

// "Đội 1 - A Nghi" → { main: "Đội 1", sub: "A Nghi" }
const splitTeam = s => {
  const [main, ...rest] = String(s || '').split(/\s+[-–]\s+/);
  return { main: (main || '').trim(), sub: rest.join(' - ').trim() };
};
// "3-1" → [3, 1]
const parseScore = s => {
  const r = String(s ?? '').match(/(\d+)\s*[-–:]\s*(\d+)/);
  return r ? [Number(r[1]), Number(r[2])] : null;
};
const FINAL_RE = /tranh nhất|nhất\s*\/\s*nhì|chung kết/i;
const THIRD_RE = /tranh ba|ba\s*\/\s*tư|hạng ba/i;
const isPlayoff = x => FINAL_RE.test(x['Vòng'] || '') || THIRD_RE.test(x['Vòng'] || '') || /xếp hạng|playoff/i.test(x['Vòng'] || '');
const isCaptain = p => p['Vai trò'] === 'Đội trưởng';

// Excel lưu giờ dưới dạng số thập phân (0.3125 = 07:30)
const fmtTime = v => {
  if (typeof v === 'number' && v >= 0 && v < 1) {
    const mins = Math.round(v * 1440);
    return String(Math.floor(mins / 60)).padStart(2, '0') + ':' + String(mins % 60).padStart(2, '0');
  }
  return v;
};
const normalizeMatch = row => {
  const o = {};
  Object.entries(row).forEach(([k, v]) => { o[k.trim()] = typeof v === 'string' ? v.trim() : v; });
  o['Giờ'] = fmtTime(o['Giờ']);
  if (typeof o['Tỉ số'] === 'number') o['Tỉ số'] = String(o['Tỉ số']);
  return o;
};

function computeStandings(matches, teams) {
  const rows = Object.fromEntries(teams.map(t => [t, { team: t, p: 0, w: 0, d: 0, l: 0, sf: 0, sa: 0 }]));
  matches.filter(x => !isPlayoff(x)).forEach(x => {
    const sc = parseScore(x['Tỉ số']);
    const a = rows[splitTeam(x['Đội 1']).main];
    const b = rows[splitTeam(x['Đội 2']).main];
    if (!sc || !a || !b) return;
    a.p++; b.p++;
    a.sf += sc[0]; a.sa += sc[1]; b.sf += sc[1]; b.sa += sc[0];
    if (sc[0] > sc[1]) { a.w++; b.l++; } else if (sc[0] < sc[1]) { b.w++; a.l++; } else { a.d++; b.d++; }
  });
  return Object.values(rows).sort((x, y) =>
    y.w - x.w || (y.sf - y.sa) - (x.sf - x.sa) || y.sf - x.sf || byNumeric(x.team, y.team));
}

function pairInfo(matches, a, b) {
  const found = matches.find(x => !isPlayoff(x) && (
    (splitTeam(x['Đội 1']).main === a && splitTeam(x['Đội 2']).main === b) ||
    (splitTeam(x['Đội 1']).main === b && splitTeam(x['Đội 2']).main === a)));
  if (!found) return null;
  const sc = parseScore(found['Tỉ số']);
  const forward = splitTeam(found['Đội 1']).main === a;
  return { match: found, score: sc ? (forward ? sc : [sc[1], sc[0]]) : null };
}

const statusClass = s =>
  /đang/i.test(s || '') ? 'status status--live' : /kết thúc|hoàn thành|xong/i.test(s || '') ? 'status status--done' : 'status';

/* ------------------------------------------------------------------ *
 *  4. HOOK
 * ------------------------------------------------------------------ */
function useCountdown(iso) {
  const target = iso ? new Date(iso).getTime() : NaN;
  const [now, setNow] = useState(() => Date.now());
  useEffect(() => {
    if (Number.isNaN(target)) return undefined;
    const id = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(id);
  }, [target]);
  if (Number.isNaN(target)) return null;
  const diff = Math.max(0, target - now);
  return {
    d: Math.floor(diff / 864e5), h: Math.floor((diff % 864e5) / 36e5),
    m: Math.floor((diff % 36e5) / 6e4), s: Math.floor((diff % 6e4) / 1e3), done: diff === 0
  };
}

function useActiveSection(ids) {
  const [active, setActive] = useState('');
  useEffect(() => {
    const els = ids.map(id => document.getElementById(id)).filter(Boolean);
    if (!('IntersectionObserver' in window)) return undefined;
    const io = new IntersectionObserver(entries => {
      entries.forEach(e => { if (e.isIntersecting) setActive(e.target.id); });
    }, { rootMargin: '-30% 0px -60% 0px' });
    els.forEach(el => io.observe(el));
    return () => io.disconnect();
  }, [ids]);
  return active;
}

/* ------------------------------------------------------------------ *
 *  5. THÀNH PHẦN GIAO DIỆN
 * ------------------------------------------------------------------ */
function Header({ active }) {
  const [open, setOpen] = useState(false);
  useEffect(() => {
    document.body.classList.toggle('no-scroll', open);
    return () => document.body.classList.remove('no-scroll');
  }, [open]);
  const close = () => setOpen(false);
  return (
    <header className="site-header">
      <div className="container header-inner">
        <a className="brand" href="#top" onClick={close}>
          <span className="brand-mark" aria-hidden="true">SI</span>
          <span className="brand-text"><strong>SI Open 2026</strong><small>MSB · Khối SI</small></span>
        </a>
        <nav id="main-nav" className={'nav' + (open ? ' is-open' : '')} aria-label="Điều hướng chính">
          {NAV.map(([id, label]) => (
            <a key={id} href={'#' + id} className={active === id ? 'is-active' : ''}
               aria-current={active === id ? 'true' : undefined} onClick={close}>{label}</a>
          ))}
          <a className="btn btn-accent nav-cta" href="#lich" onClick={close}>Tra cứu trận <ArrowUpRight size={16} aria-hidden="true" /></a>
        </nav>
        <button className="menu-toggle" type="button" aria-expanded={open} aria-controls="main-nav"
                aria-label={open ? 'Đóng menu' : 'Mở menu'} onClick={() => setOpen(o => !o)}>
          {open ? <X /> : <Menu />}
        </button>
      </div>
    </header>
  );
}

function CourtArt() {
  const still = typeof window !== 'undefined' && window.matchMedia
    && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const ballPath = 'M70 430 Q 20 250 190 96';
  return (
    <div className="court-wrap">
      <svg className="court" viewBox="0 0 260 520" role="img" aria-label="Sân pickleball nhìn từ trên xuống, quả bóng bay qua lưới">
        <rect width="260" height="520" rx="16" fill="#0c1630" />
        <rect x="30" y="40" width="200" height="440" fill="#24437f" />
        <rect x="30" y="190" width="200" height="70" fill="#ff7448" />
        <rect x="30" y="260" width="200" height="70" fill="#ff7448" />
        <g fill="none" stroke="#f5f4ef" strokeWidth="3">
          <rect x="30" y="40" width="200" height="440" />
          <line x1="30" y1="190" x2="230" y2="190" />
          <line x1="30" y1="330" x2="230" y2="330" />
          <line x1="130" y1="40" x2="130" y2="190" />
          <line x1="130" y1="330" x2="130" y2="480" />
        </g>
        <text x="130" y="232" textAnchor="middle" className="court-label">Vùng bếp</text>
        <text x="130" y="302" textAnchor="middle" className="court-label">Vùng bếp</text>
        <text x="130" y="425" textAnchor="middle" className="court-watermark">SI OPEN</text>
        <line x1="14" y1="260" x2="246" y2="260" stroke="#fff" strokeWidth="5" strokeDasharray="2 3" />
        <circle cx="14" cy="260" r="6" fill="#fff" />
        <circle cx="246" cy="260" r="6" fill="#fff" />
        <path d={ballPath} fill="none" stroke="#f5f4ef" strokeOpacity=".55" strokeWidth="2" strokeDasharray="3 7" />
        <g>
          <circle r="11" fill="#ecf24b" stroke="#111b35" strokeWidth="1.5" />
          <circle cx="-3" cy="-3" r="1.6" fill="#111b35" /><circle cx="3" cy="-2" r="1.6" fill="#111b35" />
          <circle cx="-1" cy="3" r="1.6" fill="#111b35" /><circle cx="4" cy="4" r="1.6" fill="#111b35" />
          {still
            ? <animateTransform attributeName="transform" type="translate" values="190 96" dur="1s" fill="freeze" />
            : <animateMotion dur="2.8s" repeatCount="indefinite" path={ballPath} keyPoints="0;1" keyTimes="0;1" calcMode="spline" keySplines="0.4 0 0.2 1" />}
        </g>
      </svg>
    </div>
  );
}

function Hero({ teamCount, playerCount }) {
  const cd = useCountdown(EVENT.startISO);
  return (
    <section className="hero" id="top">
      <div className="container hero-grid">
        <div className="hero-copy">
          <p className="hero-org">MSB · Strategy &amp; Innovation</p>
          <h1><span>Pickleball</span><span>SI Open</span><span>2026</span></h1>
          <p className="hero-tag">Thể thao · Kết nối · Gắn kết</p>
          <p className="hero-lead">Một buổi sáng thứ Bảy để cả Khối SI đứng chung một sân: {teamCount} đội, một thể thức vòng tròn và những trận đấu đáng nhớ.</p>
          <div className="hero-actions">
            <a className="btn btn-accent" href="#lich">Xem lịch thi đấu <ArrowUpRight size={18} aria-hidden="true" /></a>
            <a className="btn btn-ghost" href="#the-le">Đọc thể lệ</a>
          </div>
          {cd && !cd.done && (
            <div className="countdown" role="timer" aria-label="Thời gian còn lại đến giờ khai cuộc">
              {[['d', 'ngày'], ['h', 'giờ'], ['m', 'phút'], ['s', 'giây']].map(([k, label]) => (
                <div key={k}><b>{String(cd[k]).padStart(2, '0')}</b><span>{label}</span></div>
              ))}
            </div>
          )}
        </div>
        <CourtArt />
      </div>
      <div className="facts-band">
        <dl className="container facts">
          <div><span className="fact-icon"><CalendarDays size={20} aria-hidden="true" /></span><div><dt>Thời gian</dt><dd>{EVENT.dateLabel} · {EVENT.timeLabel}</dd></div></div>
          <div><span className="fact-icon"><Users size={20} aria-hidden="true" /></span><div><dt>Thể thức</dt><dd>{teamCount} đội · {playerCount} vận động viên</dd></div></div>
          <div><span className="fact-icon"><MapPin size={20} aria-hidden="true" /></span><div><dt>Địa điểm</dt><dd>{EVENT.venue}</dd></div></div>
          <div><span className="fact-icon"><Trophy size={20} aria-hidden="true" /></span><div><dt>Tổng giải thưởng</dt><dd>{EVENT.prizeTotal}</dd></div></div>
        </dl>
      </div>
    </section>
  );
}

function SectionHead({ title, lead, aside }) {
  return (
    <div className="section-head">
      <div>
        <h2>{title}</h2>
        {lead && <p className="section-lead">{lead}</p>}
      </div>
      {aside}
    </div>
  );
}

function Roster({ players }) {
  const teams = useMemo(() => {
    const map = new Map();
    players.forEach(p => { const t = p['Đội']; if (!t) return; if (!map.has(t)) map.set(t, []); map.get(t).push(p); });
    return [...map].sort((a, b) => byNumeric(a[0], b[0]));
  }, [players]);
  return (
    <section className="section" id="doi-thi">
      <div className="container">
        <SectionHead title="Bốn đội, một sân chơi chung"
          lead="Danh sách vận động viên theo thông tin Ban Tổ chức cung cấp, có thể điều chỉnh trước ngày thi đấu." />
        <div className="roster">
          {teams.map(([team, members], i) => {
            const cap = members.find(isCaptain);
            const rest = members.filter(p => p !== cap);
            return (
              <article className="team" key={team}>
                <header className="team-head">
                  <span className="team-no" aria-hidden="true">{i + 1}</span>
                  <h3>{team}</h3>
                  <span className="team-count">{members.length} VĐV</span>
                </header>
                {cap && <p className="captain"><span className="badge">Đội trưởng</span>{cap['Họ tên']}</p>}
                <ul className="team-list">
                  {rest.map((p, j) => <li key={p['Họ tên'] + j}>{p['Họ tên']}</li>)}
                </ul>
              </article>
            );
          })}
        </div>
      </div>
    </section>
  );
}

function Timeline() {
  return (
    <section className="section section--paper" id="lich-trinh">
      <div className="container">
        <SectionHead title="Từ khai mạc đến trao giải" lead="Khung giờ theo proposal tổ chức giải, kéo dài khoảng bốn tiếng rưỡi." />
        <ol className="timeline">
          {TIMELINE.map(([time, tag, title, text]) => (
            <li className="tl-item" key={time}>
              <time className="tl-time">{time}</time>
              <span className="tl-rail" aria-hidden="true" />
              <div className="tl-body">
                <span className="tag">{tag}</span>
                <h3>{title}</h3>
                <p>{text}</p>
              </div>
            </li>
          ))}
        </ol>
      </div>
    </section>
  );
}

function Side({ raw, score, other, align }) {
  const t = splitTeam(raw);
  const win = score != null && other != null && score > other;
  return (
    <div className={'side side--' + align + (win ? ' is-win' : '')}>
      <span className="side-name">{t.main || '—'}</span>
      {t.sub && <span className="side-sub">{t.sub}</span>}
    </div>
  );
}

function MatchRow({ x }) {
  const sc = parseScore(x['Tỉ số']);
  return (
    <li className="match">
      <div className="match-meta">
        <strong>{x['Sân'] || 'Chưa xếp sân'}</strong>
        <span>{[x['Mã trận'], x['Nội dung']].filter(Boolean).join(' · ')}</span>
      </div>
      <Side raw={x['Đội 1']} score={sc?.[0]} other={sc?.[1]} align="left" />
      <div className="score" aria-label={sc ? `Tỉ số ${sc[0]} - ${sc[1]}` : 'Chưa có tỉ số'}>
        {sc ? <><b className={sc[0] > sc[1] ? 'is-win' : ''}>{sc[0]}</b><i>–</i><b className={sc[1] > sc[0] ? 'is-win' : ''}>{sc[1]}</b></> : <span className="vs">vs</span>}
      </div>
      <Side raw={x['Đội 2']} score={sc?.[1]} other={sc?.[0]} align="right" />
      <span className={statusClass(x['Trạng thái'])}>{x['Trạng thái'] || 'Chưa đấu'}</span>
    </li>
  );
}

function MatchCenter({ matches, onOpenAdmin }) {
  const [query, setQuery] = useState('');
  const [event, setEvent] = useState('');
  const [round, setRound] = useState('');
  const [court, setCourt] = useState('');
  const uniq = key => [...new Set(matches.map(x => x[key]).filter(Boolean))].sort(byNumeric);
  const events = useMemo(() => uniq('Nội dung'), [matches]); // eslint-disable-line react-hooks/exhaustive-deps
  const rounds = useMemo(() => uniq('Vòng'), [matches]);     // eslint-disable-line react-hooks/exhaustive-deps
  const courts = useMemo(() => uniq('Sân'), [matches]);      // eslint-disable-line react-hooks/exhaustive-deps

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return matches.filter(x =>
      (!q || Object.values(x).join(' ').toLowerCase().includes(q)) &&
      (!event || x['Nội dung'] === event) && (!round || x['Vòng'] === round) && (!court || x['Sân'] === court));
  }, [matches, query, event, round, court]);

  const slots = useMemo(() => {
    const g = new Map();
    [...filtered]
      .sort((a, b) => String(a['Giờ']).localeCompare(String(b['Giờ'])) || byNumeric(a['Sân'] || '', b['Sân'] || ''))
      .forEach(x => { const k = x['Giờ'] || 'Chưa xếp giờ'; if (!g.has(k)) g.set(k, []); g.get(k).push(x); });
    return [...g];
  }, [filtered]);

  const hasFilter = query || event || round || court;
  const reset = () => { setQuery(''); setEvent(''); setRound(''); setCourt(''); };

  return (
    <section className="section" id="lich">
      <div className="container">
        <SectionHead title="Tìm trận của bạn trong vài giây"
          lead="Gõ tên vận động viên, tên đội hoặc số sân. Mỗi cặp đội gồm 4 trận nhỏ: 2 đôi nam, 1 đôi nữ và 1 đôi nam nữ."
          aside={<p className="count"><b>{String(matches.length).padStart(2, '0')}</b> trận<br />trong lịch</p>} />

        <div className="toolbar">
          <label className="search">
            <Search size={20} aria-hidden="true" />
            <span className="sr-only">Tìm trận đấu</span>
            <input type="search" value={query} onChange={e => setQuery(e.target.value)}
                   placeholder="Tên VĐV, tên đội, “Sân 2”…" />
          </label>
          <div className="filters">
            {events.length > 1 && (
              <select aria-label="Lọc theo nội dung" value={event} onChange={e => setEvent(e.target.value)}>
                <option value="">Tất cả nội dung</option>{events.map(v => <option key={v}>{v}</option>)}
              </select>
            )}
            {rounds.length > 1 && (
              <select aria-label="Lọc theo vòng" value={round} onChange={e => setRound(e.target.value)}>
                <option value="">Tất cả vòng</option>{rounds.map(v => <option key={v}>{v}</option>)}
              </select>
            )}
            {courts.length > 0 && (
              <div className="chips" role="group" aria-label="Lọc theo sân">
                <button type="button" className={'chip' + (!court ? ' is-on' : '')} aria-pressed={!court} onClick={() => setCourt('')}>Tất cả sân</button>
                {courts.map(c => (
                  <button type="button" key={c} className={'chip' + (court === c ? ' is-on' : '')} aria-pressed={court === c}
                          onClick={() => setCourt(court === c ? '' : c)}>{c}</button>
                ))}
              </div>
            )}
            {hasFilter && <button type="button" className="link-btn" onClick={reset}><RotateCcw size={14} aria-hidden="true" /> Xoá bộ lọc</button>}
          </div>
        </div>

        {slots.length > 0 ? (
          <div className="slots" aria-live="polite">
            {slots.map(([time, list]) => (
              <div className="slot" key={time}>
                <h3 className="slot-time"><Clock3 size={18} aria-hidden="true" />{time}</h3>
                <ul className="slot-list">{list.map((x, i) => <MatchRow x={x} key={(x['Mã trận'] || '') + i} />)}</ul>
              </div>
            ))}
          </div>
        ) : (
          <div className="empty">
            <CalendarDays size={28} aria-hidden="true" />
            <h3>{matches.length === 0 ? 'Chưa có lịch thi đấu' : 'Không tìm thấy trận phù hợp'}</h3>
            <p>{matches.length === 0
              ? 'Ban Tổ chức sẽ cập nhật lịch thi đấu trước ngày diễn ra giải.'
              : 'Thử đổi từ khoá hoặc xoá bộ lọc để xem lại toàn bộ lịch.'}</p>
            {matches.length === 0
              ? <button type="button" className="btn btn-dark" onClick={onOpenAdmin}>Nạp lịch từ Excel</button>
              : <button type="button" className="btn btn-line" onClick={reset}>Xoá bộ lọc</button>}
          </div>
        )}
      </div>
    </section>
  );
}

function Standings({ standings, teams, matches, captains }) {
  const anyPlayed = standings.some(r => r.p > 0);
  return (
    <section className="section section--paper" id="bang">
      <div className="container">
        <SectionHead title="Bảng xếp hạng vòng tròn"
          lead="Xếp hạng theo số trận thắng, sau đó đến hiệu số trận nhỏ. Bảng tự cập nhật khi Ban Tổ chức nhập tỉ số."
          aside={!anyPlayed && <p className="pill">Chưa có kết quả</p>} />
        <div className="two-col">
          <div>
            <div className="table-wrap">
              <table className="table">
                <caption className="sr-only">Bảng xếp hạng các đội</caption>
                <thead>
                  <tr>
                    <th scope="col">Hạng</th><th scope="col">Đội</th>
                    <th scope="col" className="num">Trận</th><th scope="col" className="num">Thắng</th>
                    <th scope="col" className="num">Hoà</th><th scope="col" className="num">Thua</th>
                    <th scope="col" className="num">Trận nhỏ</th><th scope="col" className="num">Hiệu số</th>
                  </tr>
                </thead>
                <tbody>
                  {standings.map((r, i) => (
                    <tr key={r.team} className={i < 2 ? 'row-top' : 'row-low'}>
                      <td className="rank">{anyPlayed ? i + 1 : '–'}</td>
                      <th scope="row"><span className="t-name">{r.team}</span>{captains[r.team] && <span className="t-sub">{captains[r.team]}</span>}</th>
                      <td className="num">{r.p}</td><td className="num">{r.w}</td><td className="num">{r.d}</td><td className="num">{r.l}</td>
                      <td className="num">{r.sf}–{r.sa}</td>
                      <td className="num">{r.sf - r.sa > 0 ? '+' : ''}{r.sf - r.sa}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <p className="legend"><i className="dot dot--top" /> Hạng 1–2 tranh Nhất/Nhì <i className="dot dot--low" /> Hạng 3–4 tranh Ba/Tư</p>
          </div>

          <div>
            <h3 className="subhead">Bảng chéo</h3>
            <p className="subnote">Đọc theo hàng: kết quả của đội ở hàng so với đội ở cột. Ô chưa có tỉ số hiện giờ và sân thi đấu.</p>
            <div className="table-wrap">
              <table className="table cross">
                <caption className="sr-only">Bảng chéo kết quả giữa các đội</caption>
                <thead>
                  <tr><td />{teams.map(t => <th scope="col" key={t}>{t}</th>)}</tr>
                </thead>
                <tbody>
                  {teams.map(a => (
                    <tr key={a}>
                      <th scope="row">{a}</th>
                      {teams.map(b => {
                        if (a === b) return <td key={b} className="self" aria-hidden="true" />;
                        const info = pairInfo(matches, a, b);
                        if (!info) return <td key={b} className="cell-empty">–</td>;
                        if (info.score) {
                          const w = info.score[0] > info.score[1];
                          return <td key={b} className={'cell-score ' + (w ? 'is-win' : 'is-loss')}>{info.score[0]}–{info.score[1]}</td>;
                        }
                        return <td key={b} className="cell-plan"><b>{info.match['Giờ'] || '—'}</b><span>{info.match['Sân'] || ''}</span></td>;
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    </section>
  );
}

function BracketMatch({ title, a, b }) {
  const pending = s => /^Hạng \d$/.test(s.name);
  return (
    <div className="bk-match">
      <p className="bk-title">{title}</p>
      {[a, b].map((s, i) => (
        <div key={i} className={'bk-row' + (pending(s) ? ' is-pending' : '') + (s.win ? ' is-win' : '')}>
          <span>{s.name}</span><b>{s.score ?? '–'}</b>
        </div>
      ))}
    </div>
  );
}

function Bracket({ matches, standings }) {
  const rr = matches.filter(x => !isPlayoff(x));
  const complete = rr.length > 0 && rr.every(x => parseScore(x['Tỉ số']));
  const seat = i => (complete && standings[i] ? standings[i].team : `Hạng ${i + 1}`);

  const build = (re, n1, n2) => {
    const x = matches.find(v => re.test(v['Vòng'] || '') && (v['Đội 1'] || v['Đội 2']));
    if (!x) return { a: { name: n1 }, b: { name: n2 }, sc: null };
    const sc = parseScore(x['Tỉ số']);
    return {
      a: { name: splitTeam(x['Đội 1']).main || n1, score: sc?.[0], win: !!sc && sc[0] > sc[1] },
      b: { name: splitTeam(x['Đội 2']).main || n2, score: sc?.[1], win: !!sc && sc[1] > sc[0] },
      sc
    };
  };
  const final = build(FINAL_RE, seat(0), seat(1));
  const third = build(THIRD_RE, seat(2), seat(3));
  const champion = final.sc && final.sc[0] !== final.sc[1] ? (final.sc[0] > final.sc[1] ? final.a.name : final.b.name) : null;

  return (
    <section className="section" id="nhanh">
      <div className="container">
        <SectionHead title="Con đường tới ngôi vô địch"
          lead="Sau vòng tròn, hạng 1–2 tranh Nhất/Nhì; hạng 3–4 tranh Ba/Tư. Tên đội sẽ tự điền khi vòng tròn kết thúc." />
        <div className="bracket">
          <BracketMatch title="Tranh Nhất/Nhì" a={final.a} b={final.b} />
          <span className="bk-link" aria-hidden="true" />
          <div className="bk-champ">
            <Trophy size={34} aria-hidden="true" />
            <strong>{champion || 'Nhà vô địch'}</strong>
            <span>{champion ? 'SI Open 2026' : 'Chờ trận tranh Nhất/Nhì'}</span>
          </div>
          <div className="bk-third"><BracketMatch title="Tranh Ba/Tư" a={third.a} b={third.b} /></div>
        </div>
      </div>
    </section>
  );
}

function Rules() {
  return (
    <section className="section section--paper" id="the-le">
      <div className="container rules-layout">
        <SectionHead title="Thể lệ thi đấu" lead="Tóm tắt theo proposal tổ chức giải. Bấm từng mục để xem chi tiết." />
        <div className="rules">
          {RULES.map((r, i) => (
            <details key={r.title} open={i === 0}>
              <summary>{r.title}</summary>
              <div className="rule-body">{r.body}</div>
            </details>
          ))}
        </div>
      </div>
    </section>
  );
}

function Awards() {
  return (
    <section className="awards" id="giai-thuong">
      <div className="container">
        <SectionHead title="Vinh danh những nỗ lực"
          lead={`Tổng giá trị giải thưởng ${EVENT.prizeTotal}, theo proposal tổ chức giải.`} />
        <div className="podium">
          {PRIZES.map(p => (
            <article className={'bar bar--' + p.cls} key={p.rank}>
              <h3>{p.rank}</h3>
              <strong>{p.amount}</strong>
              <span>{p.reward}</span>
            </article>
          ))}
        </div>
      </div>
    </section>
  );
}

function Closing() {
  return (
    <section className="closing">
      <div className="container">
        <h2>Hẹn gặp tại sân.</h2>
        <p>{EVENT.dateLabel} · {EVENT.timeLabel} · {EVENT.venue}</p>
        <a className="btn btn-accent" href="#lich">Tìm trận của bạn <ArrowUpRight size={18} aria-hidden="true" /></a>
      </div>
    </section>
  );
}

function Admin({ onClose, onApply, onReset }) {
  const [file, setFile] = useState(null);
  const [notice, setNotice] = useState('');
  const [drag, setDrag] = useState(false);

  const read = async f => {
    if (!f) return;
    if (!/\.xlsx?$/i.test(f.name)) { setNotice('Chỉ nhận file Excel .xlsx hoặc .xls.'); return; }
    if (f.size > 5 * 1024 * 1024) { setNotice('File lớn hơn 5 MB. Hãy rút gọn file rồi thử lại.'); return; }
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type: 'array' });
      const sheet = n => (wb.Sheets[n] ? XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: '' }) : []);
      const ms = sheet('Lịch thi đấu').map(normalizeMatch).filter(x => x['Đội 1'] || x['Đội 2']);
      const ps = sheet('Vận động viên').filter(p => String(p['Họ tên'] || '').trim());
      if (!ms.length && !ps.length) {
        setFile(null);
        setNotice('Không tìm thấy dữ liệu. File cần có sheet “Lịch thi đấu” hoặc “Vận động viên”.');
        return;
      }
      setFile({ matches: ms, players: ps });
      setNotice(`Đã đọc ${ms.length} trận và ${ps.length} vận động viên. Bấm “Áp dụng” để xem trên trình duyệt này.`);
    } catch {
      setFile(null);
      setNotice('Không đọc được file. Hãy kiểm tra lại định dạng .xlsx hoặc .xls.');
    }
  };

  const downloadTemplate = () => {
    const wb = XLSX.utils.book_new();
    const matchRows = [['Mã trận', 'Nội dung', 'Vòng', 'Bảng', 'Sân', 'Giờ', 'Đội 1', 'Đội 2', 'Tỉ số', 'Trạng thái'],
      ['M01', 'Đồng đội', 'Vòng tròn', 'Vòng bảng', 'Sân 1', '07:30', 'Đội 1 - A Nghi', 'Đội 2 - A Hải', '', 'Dự kiến']];
    const playerRows = [['Đội', 'Họ tên', 'Vai trò', 'Phòng ban', 'Nội dung', 'Bảng'],
      ['Đội 1', 'A Nghi', 'Đội trưởng', '', '', '']];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(matchRows), 'Lịch thi đấu');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(playerRows), 'Vận động viên');
    XLSX.writeFile(wb, 'Mau_du_lieu_SI_Open_2026.xlsx');
  };

  return (
    <section className="admin" id="btc">
      <div className="container">
        <div className="admin-head">
          <h2>Khu vực Ban Tổ chức</h2>
          <button type="button" className="icon-btn" onClick={onClose} aria-label="Đóng khu vực Ban Tổ chức"><X size={20} /></button>
        </div>
        <p className="admin-lead">Tải lên file Excel có sheet <b>“Lịch thi đấu”</b> và/hoặc <b>“Vận động viên”</b>. Cột “Vai trò” dùng giá trị <b>Đội trưởng</b>. Nhập tỉ số dạng <b>3-1</b> và đặt định dạng ô là Text để Excel không tự đổi thành ngày.</p>

        <label className={'dropzone' + (drag ? ' is-drag' : '')}
               onDragOver={e => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)}
               onDrop={e => { e.preventDefault(); setDrag(false); read(e.dataTransfer.files?.[0]); }}>
          <Upload size={26} aria-hidden="true" />
          <strong>Kéo thả file Excel vào đây hoặc bấm để chọn</strong>
          <span>.xlsx / .xls · tối đa 5 MB</span>
          <input type="file" accept=".xlsx,.xls" onChange={e => { read(e.target.files?.[0]); e.target.value = ''; }} />
        </label>

        {notice && <p className="notice" role="status">{notice}</p>}

        <div className="admin-actions">
          {file && <>
            <button type="button" className="btn btn-accent" onClick={() => { onApply(file); setFile(null); setNotice('Đã áp dụng trên trình duyệt này.'); }}><Check size={16} aria-hidden="true" /> Áp dụng</button>
            <button type="button" className="btn btn-line" onClick={() => { setFile(null); setNotice(''); }}>Huỷ</button>
          </>}
          <button type="button" className="btn btn-line" onClick={downloadTemplate}><Download size={16} aria-hidden="true" /> Tải file mẫu</button>
          <button type="button" className="btn btn-line" onClick={() => { onReset(); setNotice('Đã quay về dữ liệu mặc định.'); }}><RotateCcw size={16} aria-hidden="true" /> Về dữ liệu mặc định</button>
        </div>

        <p className="warning"><Info size={18} aria-hidden="true" /><span><b>Lưu ý:</b> đây là website tĩnh, dữ liệu nạp ở đây chỉ lưu trong trình duyệt của bạn và không hiển thị cho người xem khác. Muốn cập nhật công khai, hãy sửa dữ liệu trong mã nguồn rồi triển khai lại, hoặc kết nối cơ sở dữ liệu.</span></p>
      </div>
    </section>
  );
}

function Footer({ onOpenAdmin }) {
  return (
    <footer className="site-footer">
      <div className="container footer-inner">
        <a className="footer-brand" href="#top"><span className="brand-mark" aria-hidden="true">SI</span>SI Open 2026</a>
        <p>© 2026 MSB · Khối SI</p>
        <div className="footer-links">
          <a href="#btc" onClick={e => { e.preventDefault(); onOpenAdmin(); }}>Khu vực BTC</a>
          <a href="#top">Lên đầu trang</a>
        </div>
      </div>
    </footer>
  );
}

/* ------------------------------------------------------------------ *
 *  6. ỨNG DỤNG
 * ------------------------------------------------------------------ */
const SECTION_IDS = NAV.map(([id]) => id);

function App() {
  const [matches, setMatches] = useState(() => readSaved(KEYS.matches, DEFAULT_MATCHES));
  const [players, setPlayers] = useState(() => readSaved(KEYS.players, DEFAULT_PLAYERS));
  const [adminOpen, setAdminOpen] = useState(false);
  const active = useActiveSection(SECTION_IDS);

  const teams = useMemo(() => {
    const fromPlayers = [...new Set(players.map(p => p['Đội']).filter(Boolean))];
    const list = fromPlayers.length ? fromPlayers
      : [...new Set(matches.flatMap(x => [splitTeam(x['Đội 1']).main, splitTeam(x['Đội 2']).main]).filter(Boolean))];
    return list.sort(byNumeric);
  }, [players, matches]);
  const captains = useMemo(() => Object.fromEntries(players.filter(isCaptain).map(p => [p['Đội'], p['Họ tên']])), [players]);
  const standings = useMemo(() => computeStandings(matches, teams), [matches, teams]);

  const openAdmin = () => {
    setAdminOpen(true);
    setTimeout(() => document.getElementById('btc')?.scrollIntoView({ behavior: 'smooth' }), 60);
  };

  useEffect(() => {
    const sync = () => { if (window.location.hash === '#btc') setAdminOpen(true); };
    sync();
    window.addEventListener('hashchange', sync);
    return () => window.removeEventListener('hashchange', sync);
  }, []);

  const apply = data => {
    if (data.matches.length) { setMatches(data.matches); writeSaved(KEYS.matches, data.matches); }
    if (data.players.length) { setPlayers(data.players); writeSaved(KEYS.players, data.players); }
  };
  const reset = () => { clearSaved(); setMatches(DEFAULT_MATCHES); setPlayers(DEFAULT_PLAYERS); };

  return (
    <>
      <a className="skip-link" href="#doi-thi">Bỏ qua điều hướng</a>
      <Header active={active} />
      <main>
        <Hero teamCount={teams.length} playerCount={players.length} />
        <Roster players={players} />
        <Timeline />
        <MatchCenter matches={matches} onOpenAdmin={openAdmin} />
        <Standings standings={standings} teams={teams} matches={matches} captains={captains} />
        <Bracket matches={matches} standings={standings} />
        <Rules />
        <Awards />
        <Closing />
        {adminOpen && <Admin onClose={() => setAdminOpen(false)} onApply={apply} onReset={reset} />}
      </main>
      <Footer onOpenAdmin={openAdmin} />
    </>
  );
}

createRoot(document.getElementById('root')).render(<App />);
