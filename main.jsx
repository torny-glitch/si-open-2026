import React, { useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import {
  CalendarDays, MapPin, Search, Menu, X, ArrowUpRight, Upload, Download,
  Users, Trophy, Clock3, Info, RotateCcw, Check, HeartHandshake, Shirt, Gift, Megaphone
} from 'lucide-react';
import '@fontsource/barlow-condensed/600.css';
import '@fontsource/barlow-condensed/700.css';
import '@fontsource/barlow-condensed/800.css';
import '@fontsource/barlow-condensed/800-italic.css';
import '@fontsource/be-vietnam-pro/400.css';
import '@fontsource/be-vietnam-pro/500.css';
import '@fontsource/be-vietnam-pro/600.css';
import '@fontsource/be-vietnam-pro/700.css';
import './style.css';
import qrImage from './assets/qr.png';

/* ------------------------------------------------------------------ *
 *  1. CẤU HÌNH GIẢI — chỉ cần sửa ở đây khi có thông tin chính thức
 * ------------------------------------------------------------------ */
const EVENT = {
  name: 'MSB Pickleball SI Open 2026',
  dateLabel: 'Thứ Bảy, 07.11.2026',
  timeLabel: '07:00 – 12:00',
  venue: 'Sân pickleball Thượng Đình',
  address: '129 Nguyễn Trãi, Thanh Xuân, Hà Nội',
  startISO: '2026-11-07T07:00:00+07:00',   // bật đồng hồ đếm ngược; để '' nếu muốn tắt
  registerUrl: 'https://forms.cloud.microsoft/r/TVsfqeiqas?origin=QRCode',   // link trong mã QR
  registerDeadline: '',             // ví dụ: '17:00, 12.11.2026' — để trống thì không hiển thị
  prizeTotal: '10.000.000đ',
  playerTarget: 50,                 // số VĐV dự kiến
  teamCount: 5,                     // mỗi đội 1 đội trưởng + 9 thành viên
  courtCount: 5,
  matchTotal: 50                    // 10 lượt × 5 sân
};
const TEAM_SIZE = EVENT.playerTarget / EVENT.teamCount;   // 10 người gồm cả đội trưởng
const MATCHES_PER_TEAM = Math.round((EVENT.matchTotal * 2) / EVENT.teamCount);

const MAPS_URL = 'https://www.google.com/maps/search/?api=1&query=' +
  encodeURIComponent(`${EVENT.venue}, ${EVENT.address}`);

const pad = n => String(n).padStart(2, '0');

const LEVELS = ['A', 'B', 'C', 'D', 'E'];
const SLOTS_PER_LEVEL = EVENT.playerTarget / LEVELS.length;
const TEAMS = Array.from({ length: EVENT.teamCount }, (_, i) => `Đội ${i + 1}`);
const ROUNDS = Math.ceil(EVENT.matchTotal / EVENT.courtCount);

const NAV = [
  ['doi-thi', 'Vận động viên'],
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
  ['07:15', 'Vòng tròn', 'Thi đấu vòng tròn', `${EVENT.teamCount} đội đấu vòng tròn trên ${EVENT.courtCount} sân, tổng ${EVENT.matchTotal} trận (${ROUNDS} lượt).`],
  ['10:45', 'Xếp hạng', 'Thi đấu xếp hạng', 'Hạng 1 gặp hạng 2; hạng 3 gặp hạng 4.'],
  ['11:20', 'Lucky draw', 'Bốc thăm may mắn', 'Lucky draw vào cuối chương trình.'],
  ['11:40', 'Trao giải', 'Trao giải và chụp ảnh', 'Dự kiến kết thúc chương trình lúc 12:00.']
];

const PRIZES = [
  { rank: 'Giải Nhất', amount: '4.000.000đ', reward: 'Cúp vô địch + Huy chương vàng', cls: 'first' },
  { rank: 'Giải Nhì', amount: '3.000.000đ', reward: 'Huy chương bạc', cls: 'second' },
  { rank: 'Giải Ba', amount: '2.000.000đ', reward: 'Huy chương đồng', cls: 'third' }
];
const SIDE_PRIZES = [
  { title: 'Giải Fair Play', amount: '1.000.000đ', note: 'Vinh danh tinh thần thi đấu đẹp', Icon: HeartHandshake },
  { title: 'Giải Phong cách', amount: '', note: 'Giá trị giải sẽ được công bố', Icon: Shirt }
];

const RULES = [
  {
    title: 'Thành phần đội',
    body: (
      <p>Dự kiến {EVENT.playerTarget} vận động viên, chia thành {LEVELS.length} trình ({LEVELS[0]}–{LEVELS[LEVELS.length - 1]}),
        mỗi trình {SLOTS_PER_LEVEL} suất. Chia thành {EVENT.teamCount} đội, mỗi đội gồm 1 đội trưởng và {TEAM_SIZE - 1} thành viên.
        {EVENT.teamCount} đội trưởng là {EVENT.teamCount} vận động viên Trình A. Danh sách thành viên từng đội sẽ được công bố sau khi chia đội.</p>
    )
  },
  {
    title: 'Thể thức vòng tròn',
    body: (
      <p>{EVENT.teamCount} đội đấu vòng tròn, tổng {EVENT.matchTotal} trận; mỗi đội đấu khoảng {MATCHES_PER_TEAM} trận.
        {' '}{EVENT.courtCount} sân thi đấu song song, mỗi lượt {EVENT.courtCount} trận.</p>
    )
  },
  {
    title: 'Nội dung giao đấu',
    body: (
      <>
        <p>Thi đấu đồng đội. Mỗi lần hai đội gặp nhau có các nội dung:</p>
        <ul className="chips-static">
          <li>Đôi nam</li><li>Đôi nữ</li><li>Đôi nam nữ</li>
        </ul>
      </>
    )
  },
  {
    title: 'Cách đọc mã trận',
    body: (
      <>
        <p>Mỗi trận có một mã duy nhất. Ví dụ <b>M17</b> là trận thứ 17, thuộc <b>Lượt 4</b>, đấu ở <b>Sân 2</b>
          {' '}(công thức: số trận = (lượt − 1) × {EVENT.courtCount} + sân).</p>
        <ul className="chips-static">
          <li>Mã trận M01–M{pad(EVENT.matchTotal)}</li>
          <li>Mã sân S1–S{EVENT.courtCount}</li>
          <li>Lượt L1–L{ROUNDS}</li>
        </ul>
      </>
    )
  },
  {
    title: 'Tính điểm và xếp hạng',
    body: <p>Mỗi trận chạm 11 điểm và không có kết quả hoà. Xếp hạng theo số trận thắng, nếu bằng nhau thì xét hiệu số điểm.</p>
  },
  {
    title: 'Tranh hạng',
    body: <p>Sau vòng tròn, hạng 1 và hạng 2 tranh Nhất/Nhì; hạng 3 và hạng 4 tranh Ba/Tư.</p>
  }
];

/* ------------------------------------------------------------------ *
 *  2. DỮ LIỆU MẶC ĐỊNH (có thể thay bằng file Excel ở Khu vực BTC)
 * ------------------------------------------------------------------ */
// Khung 40 trận: M = (lượt − 1) × 5 + sân. Chưa bốc thăm nên chưa có tên đội.
const DEFAULT_MATCHES = Array.from({ length: EVENT.matchTotal }, (_, i) => ({
  'Mã trận': 'M' + pad(i + 1),
  'Lượt': Math.floor(i / EVENT.courtCount) + 1,
  'Mã sân': 'S' + ((i % EVENT.courtCount) + 1),
  'Nội dung': '', 'Vòng': 'Vòng tròn', 'Giờ': '',
  'Đội 1': '', 'Đội 2': '', 'Tỉ số': '', 'Trạng thái': 'Dự kiến'
}));

const P = (level, name, dept = '', team = '') => ({
  'Trình': level, 'Họ tên': name, 'Phòng ban': dept, 'Vai trò': team ? 'Đội trưởng' : '', 'Đội': team
});
const DEFAULT_PLAYERS = [
  // 5 đội trưởng = 5 người Trình A (số đội tạm theo thứ tự cũ, chỉnh ở đây nếu cần)
  P('A', 'Nguyễn Vũ Hải', 'DC Ops', 'Đội 2'), P('A', 'Lê Trọng Tú', 'Agile', 'Đội 4'),
  P('A', 'Phan Khánh Toàn', 'Loyalty', 'Đội 3'), P('A', 'Nguyễn Hữu Nghị', 'DigiLenO', 'Đội 1'),
  P('A', 'Hiếu DC', 'DC', 'Đội 5'),
  P('B', 'Nguyễn Mạnh Tuấn', 'Squad CMB - dự án Digi Leno'), P('B', 'Trần Trung Anh', 'DC'),
  P('B', 'Nguyễn Thanh Sơn', 'DC'), P('B', 'Nguyễn Hai Long', 'DC Ops'), P('B', 'Hà Duyên Tiến', 'UC EB'),
  P('C', 'Nghiêm An Việt', 'DigiLenO'), P('C', 'Lê Thanh Tuấn', 'PMO / dự án DigilenO'),
  P('C', 'Nguyễn Duy Minh Vũ', 'CM'), P('C', 'Đinh Anh Tú', 'DC'), P('C', 'Chu Mạnh Quang', 'AI'),
  P('D', 'Chị Thúy'), P('D', 'Phạm Thị Bích Nga', 'UC RB'), P('D', 'Hoàng Bích Phượng', 'Loyalty'),
  P('D', 'Nguyễn Hương Lan', 'DigiLenO'), P('D', 'Vũ Thị Vân Anh', 'UCSL - UCRB'),
  P('E', 'Lưu Lệ Chi', 'PGĐK'), P('E', 'Bùi Minh Phương', 'UC EB'), P('E', 'Nguyễn Quốc Minh Nhật', 'SI DC'),
  P('E', 'Nguyễn Ngọc Toàn', 'CM'), P('E', 'Hoàng Thị Thu Trang', 'Chiến lược')
];

const MATCH_COLS = ['Mã trận', 'Lượt', 'Mã sân', 'Nội dung', 'Vòng', 'Giờ', 'Đội 1', 'Đội 2', 'Tỉ số', 'Trạng thái'];
const PLAYER_COLS = ['Trình', 'Họ tên', 'Phòng ban', 'Vai trò', 'Đội'];

/* ------------------------------------------------------------------ *
 *  3. HÀM TIỆN ÍCH
 * ------------------------------------------------------------------ */
const KEYS = { matches: 'siopen_matches_v5', players: 'siopen_players_v5' };
const readSaved = (key, fallback) => {
  try { const v = localStorage.getItem(key); return v ? JSON.parse(v) : fallback; } catch { return fallback; }
};
const writeSaved = (key, value) => { try { localStorage.setItem(key, JSON.stringify(value)); } catch { /* bỏ qua */ } };
const clearSaved = () => { try { Object.values(KEYS).forEach(k => localStorage.removeItem(k)); } catch { /* bỏ qua */ } };

const byNumeric = (a, b) => String(a).localeCompare(String(b), 'vi', { numeric: true });

const splitTeam = s => {
  const [main, ...rest] = String(s || '').split(/\s+[-–]\s+/);
  return { main: (main || '').trim(), sub: rest.join(' - ').trim() };
};
const parseScore = s => {
  const r = String(s ?? '').match(/(\d+)\s*[-–:]\s*(\d+)/);
  return r ? [Number(r[1]), Number(r[2])] : null;
};
const FINAL_RE = /tranh nhất|nhất\s*\/\s*nhì|chung kết/i;
const THIRD_RE = /tranh ba|ba\s*\/\s*tư|hạng ba/i;
const isCaptain = p => p['Vai trò'] === 'Đội trưởng';
const isPlayoff = x => FINAL_RE.test(x['Vòng'] || '') || THIRD_RE.test(x['Vòng'] || '') || /xếp hạng|playoff/i.test(x['Vòng'] || '');

const courtOf = x => String(x['Mã sân'] || x['Sân'] || '').trim();
const courtLabel = c => (/^S\d+$/i.test(c) ? 'Sân ' + c.slice(1) : c);
const roundLabel = x => {
  const r = x['Lượt'];
  if (r === '' || r == null) return x['Giờ'] ? String(x['Giờ']) : 'Chưa xếp lượt';
  return /^\d+$/.test(String(r)) ? `Lượt ${r}` : String(r);
};

const fmtTime = v => {
  if (typeof v === 'number' && v >= 0 && v < 1) {
    const mins = Math.round(v * 1440);
    return pad(Math.floor(mins / 60)) + ':' + pad(mins % 60);
  }
  return v;
};
const trimRow = row => {
  const o = {};
  Object.entries(row).forEach(([k, v]) => { o[k.trim()] = typeof v === 'string' ? v.trim() : v; });
  return o;
};
const normalizeMatch = row => {
  const o = trimRow(row);
  o['Giờ'] = fmtTime(o['Giờ']);
  if (typeof o['Tỉ số'] === 'number') o['Tỉ số'] = String(o['Tỉ số']);
  if (/^L\d+$/i.test(String(o['Lượt'] ?? ''))) o['Lượt'] = Number(String(o['Lượt']).slice(1));
  return o;
};
const normalizePlayer = row => {
  const o = trimRow(row);
  o['Trình'] = String(o['Trình'] ?? '').replace(/^tr[ìi]nh\s*/i, '').trim().toUpperCase();
  return o;
};

function computeStandings(matches, teams) {
  const rows = Object.fromEntries(teams.map(t => [t, { team: t, p: 0, w: 0, l: 0, sf: 0, sa: 0 }]));
  matches.filter(x => !isPlayoff(x)).forEach(x => {
    const sc = parseScore(x['Tỉ số']);
    const a = rows[splitTeam(x['Đội 1']).main];
    const b = rows[splitTeam(x['Đội 2']).main];
    if (!sc || !a || !b || sc[0] === sc[1]) return;   // không có kết quả hoà
    a.p++; b.p++;
    a.sf += sc[0]; a.sa += sc[1]; b.sf += sc[1]; b.sa += sc[0];
    if (sc[0] > sc[1]) { a.w++; b.l++; } else { b.w++; a.l++; }
  });
  return Object.values(rows).sort((x, y) =>
    y.w - x.w || (y.sf - y.sa) - (x.sf - x.sa) || y.sf - x.sf || byNumeric(x.team, y.team));
}

// Thành tích đối đầu của đội a với đội b (hai đội có thể gặp nhau nhiều trận)
function headToHead(matches, a, b) {
  let total = 0, w = 0, l = 0;
  matches.filter(x => !isPlayoff(x)).forEach(x => {
    const t1 = splitTeam(x['Đội 1']).main, t2 = splitTeam(x['Đội 2']).main;
    if (!((t1 === a && t2 === b) || (t1 === b && t2 === a))) return;
    total++;
    const sc = parseScore(x['Tỉ số']);
    if (!sc || sc[0] === sc[1]) return;
    const mine = t1 === a ? sc[0] : sc[1];
    const theirs = t1 === a ? sc[1] : sc[0];
    if (mine > theirs) w++; else l++;
  });
  return { total, w, l, played: w + l };
}

const statusClass = s =>
  /đang/i.test(s || '') ? 'status status--live' : /kết thúc|hoàn thành|xong/i.test(s || '') ? 'status status--done' : 'status';

/* ---- Hook & thành phần giữ nguyên ---- */
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
          <a className="btn btn-accent nav-cta" href="#dang-ky" onClick={close}>Đăng ký ngay <ArrowUpRight size={16} aria-hidden="true" /></a>
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
        <rect width="260" height="520" rx="16" fill="#2ec27e" />
        <rect x="30" y="40" width="200" height="440" fill="#1456d6" />
        <rect x="30" y="190" width="200" height="70" fill="#4bc8e8" />
        <rect x="30" y="260" width="200" height="70" fill="#4bc8e8" />
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
          <circle r="11" fill="#fef027" stroke="#0a1f6e" strokeWidth="1.5" />
          <circle cx="-3" cy="-3" r="1.6" fill="#0a1f6e" /><circle cx="3" cy="-2" r="1.6" fill="#0a1f6e" />
          <circle cx="-1" cy="3" r="1.6" fill="#0a1f6e" /><circle cx="4" cy="4" r="1.6" fill="#0a1f6e" />
          {still
            ? <animateTransform attributeName="transform" type="translate" values="190 96" dur="1s" fill="freeze" />
            : <animateMotion dur="2.8s" repeatCount="indefinite" path={ballPath} keyPoints="0;1" keyTimes="0;1" calcMode="spline" keySplines="0.4 0 0.2 1" />}
        </g>
      </svg>
    </div>
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

function Timeline() {
  return (
    <section className="section section--paper" id="lich-trinh">
      <div className="container">
        <SectionHead title="Từ khai mạc đến trao giải" lead="Khung giờ dự kiến, từ 06:30 đến 12:00. Ban Tổ chức có thể điều chỉnh nhẹ trước ngày thi đấu." />
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
 *  THÀNH PHẦN MỚI / ĐÃ SỬA
 * ------------------------------------------------------------------ */
function Hero() {
  const cd = useCountdown(EVENT.startISO);
  return (
    <section className="hero" id="top">
      <div className="container hero-grid">
        <div className="hero-copy">
          <p className="hero-org">MSB · Strategy &amp; Innovation</p>
          <h1><span>Pickleball</span><span className="h1-band">SI Open 2026</span></h1>
          <p className="hero-tag">Bứt phá giới hạn – Gắn kết thành công</p>
          <p className="hero-lead">{EVENT.playerTarget} vận động viên, {EVENT.teamCount} đội, {EVENT.courtCount} sân thi đấu song song và cả Khối SI cùng đến cổ vũ. Một buổi sáng thứ Bảy thật vui.</p>
          <div className="hero-actions">
            <a className="btn btn-accent" href="#dang-ky">Đăng ký ngay <ArrowUpRight size={18} aria-hidden="true" /></a>
            <a className="btn btn-ghost" href="#lich">Lịch thi đấu</a>
          </div>
          {cd && !cd.done && (
            <div className="countdown" role="timer" aria-label="Thời gian còn lại đến giờ khai cuộc">
              {[['d', 'ngày'], ['h', 'giờ'], ['m', 'phút'], ['s', 'giây']].map(([k, label]) => (
                <div key={k}><b>{pad(cd[k])}</b><span>{label}</span></div>
              ))}
            </div>
          )}
        </div>
        <CourtArt />
      </div>
      <div className="facts-band">
        <dl className="container facts">
          <div><span className="fact-icon"><CalendarDays size={20} aria-hidden="true" /></span><div><dt>Thời gian</dt><dd>{EVENT.dateLabel}<small>{EVENT.timeLabel}</small></dd></div></div>
          <div><span className="fact-icon"><Users size={20} aria-hidden="true" /></span><div><dt>Thể thức</dt><dd>{EVENT.teamCount} đội · {EVENT.playerTarget} VĐV<small>{EVENT.courtCount} sân · {EVENT.matchTotal} trận</small></dd></div></div>
          <div><span className="fact-icon"><MapPin size={20} aria-hidden="true" /></span><div><dt>Địa điểm</dt><dd><a href={MAPS_URL} target="_blank" rel="noopener noreferrer">{EVENT.venue}</a><small>{EVENT.address}</small></dd></div></div>
          <div><span className="fact-icon"><Trophy size={20} aria-hidden="true" /></span><div><dt>Tổng giải thưởng</dt><dd>{EVENT.prizeTotal}</dd></div></div>
        </dl>
      </div>
    </section>
  );
}

function Register({ registered }) {
  const left = Math.max(0, EVENT.playerTarget - registered);
  return (
    <section className="register" id="dang-ky">
      <div className="container register-grid">
        <div>
          <h2>Quét QR, giữ ngay suất thi đấu</h2>
          <p>Còn <b>{left}</b> trong {EVENT.playerTarget} suất. Anh/chị/em đăng ký sớm nhé!</p>
          {EVENT.registerDeadline && <p className="register-deadline">Hạn chót đăng ký: <b>{EVENT.registerDeadline}</b></p>}
          <div className="register-actions">
            <a className="btn btn-dark" href={EVENT.registerUrl} target="_blank" rel="noopener noreferrer">Đăng ký ngay <ArrowUpRight size={18} aria-hidden="true" /></a>
            <a className="btn btn-line" href="#doi-thi">Xem danh sách đã đăng ký</a>
          </div>
        </div>
        <figure className="qr">
          <img src={qrImage} width="260" height="255" alt="Mã QR đăng ký tham gia giải Pickleball Khối SI 2026" />
          <figcaption>Quét bằng camera điện thoại</figcaption>
        </figure>
      </div>
    </section>
  );
}

function Cheer() {
  return (
    <section className="cheer" id="co-vu">
      <div className="container cheer-grid">
        <div>
          <h2>Đến thật đông, cổ vũ thật to!</h2>
          <p>Giải của cả Khối SI. Ai thi đấu thì đăng ký, ai chưa thi đấu thì đến sân cổ vũ đồng đội. Một tràng pháo tay cũng là cả một trận thắng.</p>
        </div>
        <div className="cheer-tiles">
          <div className="cheer-tile"><Users size={28} aria-hidden="true" /><div><h3>Thi đấu</h3><p>Quét mã QR ở phía trên để đăng ký suất thi đấu.</p></div></div>
          <div className="cheer-tile"><Megaphone size={28} aria-hidden="true" /><div><h3>Cổ vũ</h3><p>Đến {EVENT.venue}, {EVENT.dateLabel}, {EVENT.timeLabel} để tiếp lửa cho đồng đội.</p></div></div>
        </div>
      </div>
    </section>
  );
}

function Players({ players }) {
  const { groups, unassigned } = useMemo(() => {
    const g = Object.fromEntries(LEVELS.map(l => [l, []]));
    const extra = [];
    players.forEach(p => { (g[p['Trình']] || extra).push(p); });
    return { groups: g, unassigned: extra };
  }, [players]);
  const registered = players.length;
  const percent = Math.min(100, Math.round((registered / EVENT.playerTarget) * 100));

  return (
    <section className="section" id="doi-thi">
      <div className="container">
        <SectionHead title="Vận động viên theo trình"
          lead={`Dự kiến ${EVENT.playerTarget} vận động viên chia thành ${LEVELS.length} trình, mỗi trình ${SLOTS_PER_LEVEL} suất. Ô trống là những suất còn mở.`}
          aside={(
            <div className="progress">
              <p><b>{registered}</b><span>/{EVENT.playerTarget} VĐV đã có tên</span></p>
              <div className="progress-track" role="progressbar" aria-valuemin={0} aria-valuemax={EVENT.playerTarget}
                   aria-valuenow={registered} aria-label="Số vận động viên đã đăng ký"><i style={{ width: percent + '%' }} /></div>
              <small>Còn {Math.max(0, EVENT.playerTarget - registered)} suất trống</small>
            </div>
          )} />

        <div className="levels">
          {LEVELS.map(l => {
            const members = groups[l];
            const empty = Math.max(0, SLOTS_PER_LEVEL - members.length);
            return (
              <article className="level" key={l}>
                <header className="level-head">
                  <span className="level-letter" aria-hidden="true">{l}</span>
                  <h3>Trình {l}</h3>
                  <span className="level-count">{members.length}/{SLOTS_PER_LEVEL}</span>
                </header>
                <ol className="level-list">
                  {members.map((p, i) => (
                    <li className="is-filled" key={p['Họ tên'] + i}>
                      <span className="pname">{p['Họ tên']}</span>
                      {p['Phòng ban'] && <span className="pdept" title={p['Phòng ban']}>{p['Phòng ban']}</span>}
                    </li>
                  ))}
                  {Array.from({ length: empty }, (_, i) => <li className="is-empty" key={'e' + i}><span>Còn trống</span></li>)}
                </ol>
              </article>
            );
          })}
          {unassigned.length > 0 && (
            <article className="level">
              <header className="level-head"><h3>Chưa phân trình</h3><span className="level-count">{unassigned.length}</span></header>
              <ol className="level-list">
                {unassigned.map((p, i) => <li className="is-filled" key={i}><span className="pname">{p['Họ tên']}</span>{p['Phòng ban'] && <span className="pdept">{p['Phòng ban']}</span>}</li>)}
              </ol>
            </article>
          )}
        </div>

        <div className="teams-strip">
          <div>
            <h3 className="subhead">Chia đội</h3>
            <p className="subnote">
              {EVENT.teamCount} đội, mỗi đội gồm 1 đội trưởng và {TEAM_SIZE - 1} thành viên. Đội trưởng là {EVENT.teamCount} vận động viên Trình A;
              thành viên từng đội sẽ được công bố sau khi chia đội.
            </p>
          </div>
          <div className="team-cards">
            {TEAMS.map((t, i) => {
              const all = players.filter(p => p['Đội'] === t);
              const cap = all.find(isCaptain);
              const members = all.filter(p => p !== cap);
              return (
                <article className="tcard" key={t}>
                  <header><span className="team-no" aria-hidden="true">{i + 1}</span><h4>{t}</h4></header>
                  <p className="tcap">
                    <span className="badge">Đội trưởng</span>
                    {cap ? cap['Họ tên'] : 'Chưa xác định'}
                    {cap && cap['Phòng ban'] && <small>{cap['Phòng ban']}</small>}
                  </p>
                  <p className="tmembers"><b>{members.length}/{TEAM_SIZE - 1}</b> thành viên{members.length === 0 ? ' · chờ chia đội' : ''}</p>
                  {members.length > 0 && <ul className="tlist">{members.map((m, k) => <li key={m['Họ tên'] + k}>{m['Họ tên']}</li>)}</ul>}
                </article>
              );
            })}
          </div>
        </div>
      </div>
    </section>
  );
}

function MatchCard({ x }) {
  const sc = parseScore(x['Tỉ số']);
  const a = splitTeam(x['Đội 1']).main;
  const b = splitTeam(x['Đội 2']).main;
  const pending = !a && !b;
  const rows = [[a, sc?.[0], sc?.[1]], [b, sc?.[1], sc?.[0]]];
  return (
    <li className="mcard">
      <div className="mcard-top">
        <strong className="mcode">{x['Mã trận'] || '—'}</strong>
        <span className="mcourt">{courtLabel(courtOf(x)) || 'Chưa xếp sân'}</span>
      </div>
      {pending ? (
        <p className="mcard-wait">Chờ bốc thăm</p>
      ) : (
        <div className="mcard-teams">
          {rows.map(([name, mine, theirs], i) => (
            <div key={i} className={'mteam' + (mine != null && mine > theirs ? ' is-win' : '')}>
              <span>{name || '—'}</span><b>{mine ?? ''}</b>
            </div>
          ))}
        </div>
      )}
      <div className="mcard-foot">
        <span className="mnote">{x['Nội dung'] || ''}</span>
        <span className={statusClass(x['Trạng thái'])}>{x['Trạng thái'] || 'Chưa đấu'}</span>
      </div>
    </li>
  );
}

function MatchCenter({ matches, onOpenAdmin }) {
  const [query, setQuery] = useState('');
  const [event, setEvent] = useState('');
  const [round, setRound] = useState('');
  const [court, setCourt] = useState('');
  const uniq = list => [...new Set(list.filter(Boolean))].sort(byNumeric);
  const events = useMemo(() => uniq(matches.map(x => x['Nội dung'])), [matches]);
  const rounds = useMemo(() => uniq(matches.map(x => x['Vòng'])), [matches]);
  const courts = useMemo(() => uniq(matches.map(courtOf)), [matches]);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return matches.filter(x => {
      const hay = [...Object.values(x), roundLabel(x), courtLabel(courtOf(x)), x['Lượt'] ? 'L' + x['Lượt'] : ''].join(' ').toLowerCase();
      return (!q || hay.includes(q)) && (!event || x['Nội dung'] === event) && (!round || x['Vòng'] === round) && (!court || courtOf(x) === court);
    });
  }, [matches, query, event, round, court]);

  const slots = useMemo(() => {
    const g = new Map();
    [...filtered]
      .sort((a, b) => (Number(a['Lượt']) || 999) - (Number(b['Lượt']) || 999)
        || String(a['Giờ']).localeCompare(String(b['Giờ'])) || byNumeric(courtOf(a), courtOf(b)))
      .forEach(x => { const k = roundLabel(x); if (!g.has(k)) g.set(k, []); g.get(k).push(x); });
    return [...g];
  }, [filtered]);

  const hasFilter = query || event || round || court;
  const reset = () => { setQuery(''); setEvent(''); setRound(''); setCourt(''); };

  return (
    <section className="section" id="lich">
      <div className="container">
        <SectionHead title="Tìm trận của bạn trong vài giây"
          lead={`Gõ mã trận (M17), mã sân (S2), lượt (L4) hoặc tên đội. ${EVENT.courtCount} sân thi đấu song song, mỗi lượt ${EVENT.courtCount} trận.`}
          aside={<p className="count"><b>{pad(matches.length)}</b> trận<br />trong lịch</p>} />

        <div className="toolbar">
          <label className="search">
            <Search size={20} aria-hidden="true" />
            <span className="sr-only">Tìm trận đấu</span>
            <input type="search" value={query} onChange={e => setQuery(e.target.value)}
                   placeholder="Mã trận, sân, lượt hoặc tên đội…" />
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
                          onClick={() => setCourt(court === c ? '' : c)}>{courtLabel(c)}</button>
                ))}
              </div>
            )}
            {hasFilter && <button type="button" className="link-btn" onClick={reset}><RotateCcw size={14} aria-hidden="true" /> Xoá bộ lọc</button>}
          </div>
        </div>

        {slots.length > 0 ? (
          <div className="slots" aria-live="polite">
            {slots.map(([label, list]) => {
              const time = list.map(x => x['Giờ']).find(Boolean);
              return (
                <div className="slot" key={label}>
                  <h3 className="slot-head">
                    <span className="slot-title">{label}</span>
                    <span className="slot-time"><Clock3 size={15} aria-hidden="true" />{time || 'Giờ sẽ công bố'}</span>
                  </h3>
                  <ul className="slot-list">{list.map((x, i) => <MatchCard x={x} key={(x['Mã trận'] || '') + i} />)}</ul>
                </div>
              );
            })}
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

function Standings({ standings, teams, matches }) {
  const anyPlayed = standings.some(r => r.p > 0);
  return (
    <section className="section section--paper" id="bang">
      <div className="container">
        <SectionHead title="Bảng xếp hạng vòng tròn"
          lead="Không có kết quả hoà. Xếp hạng theo số trận thắng, sau đó đến hiệu số điểm. Bảng tự cập nhật khi Ban Tổ chức nhập tỉ số."
          aside={!anyPlayed && <p className="pill">Chưa có kết quả</p>} />
        <div className="two-col">
          <div>
            <div className="table-wrap">
              <table className="table">
                <caption className="sr-only">Bảng xếp hạng các đội</caption>
                <thead>
                  <tr>
                    <th scope="col">Hạng</th><th scope="col">Đội</th>
                    <th scope="col" className="num">Đã đấu</th><th scope="col" className="num">Thắng</th>
                    <th scope="col" className="num">Thua</th>
                    <th scope="col" className="num">Điểm ghi–thủng</th><th scope="col" className="num">Hiệu số</th>
                  </tr>
                </thead>
                <tbody>
                  {standings.map((r, i) => (
                    <tr key={r.team} className={i < 2 ? 'row-top' : i < 4 ? 'row-low' : 'row-out'}>
                      <td className="rank">{anyPlayed ? i + 1 : '–'}</td>
                      <th scope="row"><span className="t-name">{r.team}</span></th>
                      <td className="num">{r.p}<span className="of">/{MATCHES_PER_TEAM}</span></td>
                      <td className="num">{r.w}</td><td className="num">{r.l}</td>
                      <td className="num">{r.sf}–{r.sa}</td>
                      <td className="num">{r.sf - r.sa > 0 ? '+' : ''}{r.sf - r.sa}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <p className="legend"><i className="dot dot--top" /> Hạng 1–2 tranh Nhất/Nhì <i className="dot dot--low" /> Hạng 3–4 tranh Ba/Tư <i className="dot dot--out" /> Hạng 5 dừng ở vòng tròn</p>
          </div>

          <div>
            <h3 className="subhead">Bảng chéo</h3>
            <p className="subnote">Đọc theo hàng: số trận thắng–thua của đội ở hàng khi gặp đội ở cột. Dòng nhỏ là tổng số trận hai đội gặp nhau trong lịch.</p>
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
                        const h = headToHead(matches, a, b);
                        if (!h.total) return <td key={b} className="cell-empty">–</td>;
                        const cls = h.played === 0 ? 'cell-plan' : 'cell-score ' + (h.w >= h.l ? 'is-win' : 'is-loss');
                        return (
                          <td key={b} className={cls}>
                            <b>{h.played ? `${h.w}–${h.l}` : '–'}</b><span>{h.total} trận</span>
                          </td>
                        );
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
        <SectionHead title="Thể lệ thi đấu" lead="Tóm tắt thể thức dự kiến của giải. Bấm từng mục để xem chi tiết." />
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
          lead={`Tổng giá trị giải thưởng ${EVENT.prizeTotal}, gồm ba giải thứ hạng, các giải phụ và lucky draw cuối chương trình.`} />
        <div className="podium">
          {PRIZES.map(p => (
            <article className={'bar bar--' + p.cls} key={p.rank}>
              <h3>{p.rank}</h3>
              <strong>{p.amount}</strong>
              <span>{p.reward}</span>
            </article>
          ))}
        </div>
        <h3 className="extras-title">Giải phụ và lucky draw</h3>
        <div className="extras">
          {SIDE_PRIZES.map(({ title, amount, note, Icon }) => (
            <article className="extra" key={title}>
              <Icon size={30} aria-hidden="true" />
              <div><h4>{title}</h4>{amount && <strong>{amount}</strong>}<span>{note}</span></div>
            </article>
          ))}
          <article className="extra extra--lucky">
            <Gift size={30} aria-hidden="true" />
            <div><h4>Lucky draw</h4><strong>Cuối chương trình</strong><span>Bốc thăm may mắn, danh sách quà tặng sẽ được công bố sau.</span></div>
          </article>
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
        <p>{EVENT.dateLabel} · {EVENT.timeLabel}<br />
          <a href={MAPS_URL} target="_blank" rel="noopener noreferrer">{EVENT.venue}, {EVENT.address}</a></p>
        <div className="closing-actions">
          <a className="btn btn-accent" href="#dang-ky">Đăng ký ngay <ArrowUpRight size={18} aria-hidden="true" /></a>
          <a className="btn btn-ghost" href="#lich">Tìm trận của bạn</a>
        </div>
      </div>
    </section>
  );
}

function Admin({ matches, players, onClose, onApply, onReset }) {
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
      const ms = sheet('Lịch thi đấu').map(normalizeMatch).filter(x => x['Mã trận'] || x['Đội 1'] || x['Đội 2']);
      const ps = sheet('Vận động viên').map(normalizePlayer).filter(p => String(p['Họ tên'] || '').trim());
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

  // File tải về chứa dữ liệu đang hiển thị, dùng luôn làm mẫu để điền tiếp
  const download = () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(matches, { header: MATCH_COLS }), 'Lịch thi đấu');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(players, { header: PLAYER_COLS }), 'Vận động viên');
    XLSX.writeFile(wb, 'SI_Open_2026_du_lieu.xlsx');
  };

  return (
    <section className="admin" id="btc">
      <div className="container">
        <div className="admin-head">
          <h2>Khu vực Ban Tổ chức</h2>
          <button type="button" className="icon-btn" onClick={onClose} aria-label="Đóng khu vực Ban Tổ chức"><X size={20} /></button>
        </div>
        <p className="admin-lead">Tải file mẫu về, điền tiếp rồi tải lên. Sheet <b>“Vận động viên”</b> gồm cột Trình (A–E), Họ tên, Phòng ban, Vai trò (ghi <b>Đội trưởng</b> cho 5 đội trưởng) và Đội (Đội 1–Đội 5, để trống tới khi chia đội).
          Sheet <b>“Lịch thi đấu”</b> đã có sẵn {EVENT.matchTotal} mã trận, lượt và sân; chỉ cần điền Đội 1, Đội 2, Giờ, Tỉ số, Trạng thái.
          Nhập tỉ số dạng <b>11-8</b> và đặt định dạng ô là Text để Excel không tự đổi thành ngày.</p>

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
          <button type="button" className="btn btn-line" onClick={download}><Download size={16} aria-hidden="true" /> Tải file mẫu (dữ liệu hiện tại)</button>
          <button type="button" className="btn btn-line" onClick={() => { onReset(); setNotice('Đã quay về dữ liệu mặc định.'); }}><RotateCcw size={16} aria-hidden="true" /> Về dữ liệu mặc định</button>
        </div>

        <p className="warning"><Info size={18} aria-hidden="true" /><span><b>Lưu ý:</b> đây là website tĩnh, dữ liệu nạp ở đây chỉ lưu trong trình duyệt của bạn và không hiển thị cho người xem khác. Muốn cập nhật công khai, hãy sửa dữ liệu trong mã nguồn rồi triển khai lại, hoặc kết nối cơ sở dữ liệu.</span></p>
      </div>
    </section>
  );
}

/* ------------------------------------------------------------------ *
 *  ỨNG DỤNG
 * ------------------------------------------------------------------ */
const SECTION_IDS = NAV.map(([id]) => id);

function App() {
  const [matches, setMatches] = useState(() => readSaved(KEYS.matches, DEFAULT_MATCHES));
  const [players, setPlayers] = useState(() => readSaved(KEYS.players, DEFAULT_PLAYERS));
  const [adminOpen, setAdminOpen] = useState(false);
  const active = useActiveSection(SECTION_IDS);

  const teams = useMemo(() => {
    const set = new Set(TEAMS);
    matches.forEach(x => [x['Đội 1'], x['Đội 2']].forEach(t => { const n = splitTeam(t).main; if (n) set.add(n); }));
    return [...set].sort(byNumeric);
  }, [matches]);
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
        <Hero />
        <Register registered={players.length} />
        <Players players={players} />
        <Timeline />
        <MatchCenter matches={matches} onOpenAdmin={openAdmin} />
        <Standings standings={standings} teams={teams} matches={matches} />
        <Bracket matches={matches} standings={standings} />
        <Rules />
        <Awards />
        <Cheer />
        <Closing />
        {adminOpen && <Admin matches={matches} players={players} onClose={() => setAdminOpen(false)} onApply={apply} onReset={reset} />}
      </main>
      <Footer onOpenAdmin={openAdmin} />
    </>
  );
}

createRoot(document.getElementById('root')).render(<App />);
