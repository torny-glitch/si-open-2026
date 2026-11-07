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
 *  1. CẤU HÌNH GIẢI
 * ------------------------------------------------------------------ */
const EVENT = {
  name: 'MSB Pickleball SI Open 2026',
  dateLabel: 'Thứ Bảy, 07.11.2026',
  timeLabel: '07:00 – 12:00',
  venue: 'Sân pickleball Thượng Đình',
  address: '129 Nguyễn Trãi, Thanh Xuân, Hà Nội',
  startISO: '2026-11-07T07:00:00+07:00',   // để '' để tắt đồng hồ đếm ngược
  registerUrl: 'https://forms.cloud.microsoft/r/TVsfqeiqas?origin=QRCode',
  prizeTotal: '10.000.000đ',
  teamCount: 5,
  guestTeam: 5,                            // Đội 5 là đội khách mời
  teamSize: 8,                             // đội nội bộ
  guestSize: 11,                           // đội khách mời
  courtCount: 5
};

const MAPS_URL = 'https://www.google.com/maps/search/?api=1&query=' +
  encodeURIComponent(`${EVENT.venue}, ${EVENT.address}`);

const LEVELS = ['A', 'B', 'C', 'D'];
const TEAMS = [1, 2, 3, 4, 5];
const PH_RR = 'Vòng tròn';
const PH_FINAL = 'Chung kết';
const PH_THIRD = 'Tranh 3–4';
// Cách ghi cặp đấu: 'A-1' là đôi trình A của Đội 1. Đổi thành `${lvl}${team}` nếu muốn 'A1'.
const LABEL = (lvl, team) => `${lvl}-${team}`;
const teamName = n => `Đội ${n}`;
const isGuest = n => n === EVENT.guestTeam;
const encKey = (a, b) => [Math.min(a, b), Math.max(a, b)].join('-');
const pad = n => String(n).padStart(2, '0');
const byNumeric = (a, b) => String(a).localeCompare(String(b), 'vi', { numeric: true });

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
  ['01', 'Chuẩn bị', 'Chuẩn bị hậu cần', 'Ban Tổ chức chuẩn bị sân bãi và dụng cụ.'],
  ['02', 'Khai mạc', 'Check-in, khai mạc và phổ biến thể lệ', 'Ổn định đội hình, phổ biến quy định thi đấu.'],
  ['03', 'Vòng tròn', 'Thi đấu vòng tròn', '40 trận, 4 hạng mỗi cặp đội.'],
  ['04', 'Tranh 3–4', 'Tranh giải Ba', 'Hạng 3–4 đấu tranh giải Ba.'],
  ['05', 'Chung kết', 'Chung kết', 'Hạng 1–2 đấu chung kết.'],
  ['06', 'Trao giải', 'Trao giải và lucky draw', 'Kết thúc chương trình.']
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
    title: 'Đối tượng và thể thức chung',
    body: (
      <>
        <ul className="rule-list">
          <li>5 đội thi đấu đồng đội theo thể thức vòng tròn tính điểm, mỗi cặp đội gặp nhau 1 lượt.</li>
          <li>Tổng cộng 10 cuộc đối đầu, 40 trận vòng tròn.</li>
          <li>VĐV được xếp 4 trình A, B, C, D và thi đấu theo 2 hạng độc lập:
            <ul className="rule-sub">
              <li><b>Hạng Nâng cao (AB):</b> trình A và B</li>
              <li><b>Hạng Phong trào (CD):</b> trình C và D</li>
            </ul>
          </li>
        </ul>
      </>
    )
  },
  {
    title: 'Cơ cấu mỗi cuộc đối đầu',
    body: (
      <>
        <p className="rule-label">Mỗi cuộc đối đầu gồm 4 trận đánh đôi:</p>
        <ul className="rule-list">
          <li><b>AB-1 và AB-2:</b> 2 trận hạng Nâng cao</li>
          <li><b>CD-1 và CD-2:</b> 2 trận hạng Phong trào</li>
        </ul>
        <p className="rule-label">Tổng tỉ số tối đa của một cuộc đối đầu: 4–0, 3–1 hoặc 2–2.</p>
      </>
    )
  },
  {
    title: 'Ghép cặp và line-up',
    body: (
      <>
        <ul className="rule-list">
          <li>Trong mỗi hạng, đội tự sắp đôi từ các VĐV thuộc hạng đó. Hạng AB gồm trình A và B, hạng CD gồm trình C và D.</li>
          <li>Trong cùng một khung giờ, mỗi VĐV chỉ đăng ký thi đấu tối đa 1 trận.</li>
          <li>Đội trưởng nộp kín danh sách 4 đôi cho Ban Tổ chức trước mỗi cuộc đối đầu, để đảm bảo tính bất ngờ chiến thuật.</li>
          <li>Đội trưởng các đội 1–4 là VĐV trình A. Đội 5 do Trần Khắc Hiếu làm đội trưởng.</li>
          <li>Đội 5 (khách mời) có 11 người và thi đấu theo cùng lịch với các đội khác.</li>
        </ul>
      </>
    )
  },
  {
    title: 'Luật thi đấu và điểm số',
    body: (
      <>
        <ul className="rule-list">
          <li><b>Vòng tròn:</b> 1 set chạm 11 điểm, đổi sân ở điểm 6.</li>
          <li><b>Chung kết và tranh 3–4:</b> 1 set chạm 15 điểm, đổi sân ở điểm 8.</li>
          <li><b>Áp dụng Traditional Scoring:</b> chỉ bên đang giao bóng mới được tính điểm. Bên nhận giao bóng thắng pha bóng chỉ được giao bóng, không được cộng điểm.</li>
        </ul>
      </>
    )
  },
  {
    title: 'Tính điểm xếp hạng đội',
    body: (
      <>
        <p className="rule-label">Điểm của mỗi cuộc đối đầu:</p>
        <ul className="rule-list">
          <li>Thắng 4–0 hoặc 3–1: <b>3 điểm</b></li>
          <li>Hoà 2–2: <b>1 điểm</b> cho mỗi đội</li>
          <li>Thua 1–3 hoặc 0–4: <b>0 điểm</b></li>
        </ul>
        <p className="rule-label">Khi bằng điểm, xét theo thứ tự:</p>
        <ol className="rule-list rule-list--num">
          <li>Hiệu số trận thắng – thua</li>
          <li>Hiệu số điểm thắng – thua</li>
          <li>Kết quả đối đầu trực tiếp</li>
        </ol>
      </>
    )
  },
  {
    title: 'Tranh hạng',
    body: (
      <ul className="rule-list">
        <li>Hạng 1–2 đấu chung kết.</li>
        <li>Hạng 3–4 đấu tranh giải Ba.</li>
        <li>Đội hạng 5 dừng bước.</li>
      </ul>
    )
  }
];

/* ------------------------------------------------------------------ *
 *  2. DỮ LIỆU
 * ------------------------------------------------------------------ */
const KEYS = { matches: 'siopen_v10_matches', players: 'siopen_v10_players' };

// Danh sách 42 VĐV theo bảng chia đội (Đội 1–5) và bảng chia trình (A–D).
// Đội trưởng là VĐV trình A của mỗi đội nội bộ.
const P = (lvl, name, dept, team, captain = false) => ({ lvl, name, dept, team, captain });
const DEFAULT_PLAYERS = [
  P('A', 'Nguyễn Hữu Nghị', 'DigiLenO', 1, true), P('A', 'Phạm Hồng Phúc', 'DigiLenO', 1),
  P('B', 'Nguyễn Mạnh Tuấn', 'DigiLenO', 1), P('B', 'Nghiêm An Việt', 'DigiLenO', 1),
  P('C', 'Nguyễn Hương Lan', 'DigiLenO', 1), P('C', 'Vũ Thị Vân Anh', 'UCSL - UCRB', 1),
  P('D', 'Lê Cường', 'DC', 1), P('D', 'Nguyễn Quốc Minh Nhật', 'DC', 1),

  P('A', 'Trần Trung Hiếu', 'DC', 2, true), P('A', 'Nguyễn Vũ Hải', 'DC', 2),
  P('B', 'Đinh Anh Tú', 'DC', 2), P('B', 'Nguyễn Thanh Sơn', 'DC', 2),
  P('C', 'Trần Trung Anh', 'DC', 2), P('C', 'Ngô Hoàng Việt', 'DC', 2),
  P('D', 'Nguyễn Đăng Thiên', 'DC', 2), P('D', 'Nguyễn Lê Tới', 'DC', 2),

  P('A', 'Phan Khánh Toàn', 'Loyalty', 3, true), P('A', 'Hà Duyên Tiến', 'UC EB', 3),
  P('B', 'Nguyễn Duy Minh Vũ', 'CM', 3), P('B', 'Lê Thanh Tuấn', 'DigiLenO', 3),
  P('C', 'Hà Thị Phương Thúy', 'Loyalty', 3), P('C', 'Hoàng Bích Phượng', 'Loyalty', 3),
  P('D', 'Nguyễn Thị Loan', 'AI', 3), P('D', 'Nguyễn Ngọc Toàn', 'CM', 3),

  P('A', 'Lê Trọng Tú', 'Agile', 4, true), P('A', 'Trần Anh Kiên', '', 4),
  P('B', 'Chu Mạnh Quang', 'Agile', 4), P('B', 'Nguyễn Hai Long', 'DC Ops', 4),
  P('C', 'Nguyễn Thanh Tùng', 'Khách mời', 4), P('C', 'Nguyễn Thanh Hà', 'DigiLenO', 4),
  P('D', 'Nguyễn Văn Đạt', 'UC RB', 4), P('D', 'Dương Thùy Dung', 'UC RB', 4),

  P('A', 'Vũ Ngọc Kỳ', 'Khách mời', 5), P('A', 'Nguyễn Đức Thắng', 'Khách mời', 5),
  P('B', 'Vũ Hồng Minh', 'Khách mời', 5), P('B', 'Trần Khắc Hiếu', 'Khách mời', 5),
  P('B', 'Dương Hoàng Long', 'Khách mời', 5, true), P('B', 'Phạm Văn Thắng', 'Khách mời', 5),
  P('C', 'Nguyễn Hoàng Thái', 'Khách mời', 5), P('C', 'Hoàng Ngọc Thắng', 'Khách mời', 5),
  P('D', 'Văn Minh Phương', 'Khách mời', 5),
  P('D', 'Nguyễn Thị Nhàn', 'Khách mời', 5), P('D', 'Bùi Thế Anh', 'Khách mời', 5)
];

// Lịch 40 trận vòng tròn: [giờ, sân, hạng, số trận trong hạng, đội 1, đội 2]. Mỗi cặp đội có AB-1, AB-2, CD-1, CD-2.
// Đội tự sắp đôi trong hạng AB hoặc CD. Tranh 3–4 lúc 11:00 và chung kết lúc 11:30, mỗi trận 4 hạng trên sân 16–19.
const SCHEDULE = [
  ['07:30', 12, 'AB', 1, 1, 2],
  ['07:30', 15, 'AB', 2, 1, 2],
  ['07:30', 16, 'CD', 1, 1, 2],
  ['07:30', 17, 'CD', 2, 1, 2],
  ['07:30', 18, 'AB', 1, 3, 4],
  ['07:30', 19, 'AB', 2, 3, 4],
  ['08:00', 12, 'CD', 1, 3, 4],
  ['08:00', 15, 'CD', 2, 3, 4],
  ['08:00', 16, 'CD', 1, 1, 5],
  ['08:00', 17, 'CD', 2, 1, 5],
  ['08:00', 18, 'AB', 1, 1, 5],
  ['08:00', 19, 'AB', 2, 1, 5],
  ['08:30', 12, 'AB', 1, 2, 3],
  ['08:30', 15, 'AB', 2, 2, 3],
  ['08:30', 16, 'CD', 1, 2, 3],
  ['08:30', 17, 'CD', 2, 2, 3],
  ['08:30', 18, 'AB', 1, 4, 5],
  ['08:30', 19, 'AB', 2, 4, 5],
  ['09:00', 12, 'CD', 1, 4, 5],
  ['09:00', 15, 'CD', 2, 4, 5],
  ['09:00', 16, 'AB', 1, 1, 3],
  ['09:00', 17, 'AB', 2, 1, 3],
  ['09:00', 18, 'CD', 1, 1, 3],
  ['09:00', 19, 'CD', 2, 1, 3],
  ['09:30', 12, 'AB', 1, 2, 4],
  ['09:30', 15, 'AB', 2, 2, 4],
  ['09:30', 16, 'CD', 1, 2, 4],
  ['09:30', 17, 'CD', 2, 2, 4],
  ['09:30', 18, 'AB', 1, 2, 5],
  ['09:30', 19, 'AB', 2, 2, 5],
  ['10:00', 12, 'CD', 1, 2, 5],
  ['10:00', 15, 'CD', 2, 2, 5],
  ['10:00', 16, 'AB', 1, 1, 4],
  ['10:00', 17, 'AB', 2, 1, 4],
  ['10:00', 18, 'CD', 1, 1, 4],
  ['10:00', 19, 'CD', 2, 1, 4],
  ['10:30', 12, 'AB', 1, 3, 5],
  ['10:30', 15, 'AB', 2, 3, 5],
  ['10:30', 16, 'CD', 1, 3, 5],
  ['10:30', 17, 'CD', 2, 3, 5]
];
// Kết quả vòng tròn: [đội 1, đội 2, hạng, số trận, tỉ số (đội 1 – đội 2), tên VĐV đội 1, tên VĐV đội 2, vòng]
const RESULTS = [
  [1, 2, 'AB', 1, '10-11', 'Nghị - Phúc', 'Hiếu - Sơn', 1],
  [1, 2, 'AB', 2, '4-11', 'Tuấn - Việt', 'Hải - Tú', 1],
  [1, 2, 'CD', 1, '4-11', 'Cường - Lan', 'Tới - Việt', 1],
  [1, 2, 'CD', 2, '6-11', 'Cường - Vân Anh', 'Thiên - Trung Anh', 1],
  [3, 4, 'AB', 1, '6-11', 'Toàn - Tiến', 'Tú - Quang', 1],
  [3, 4, 'AB', 2, '3-11', 'Tuấn - Vũ', 'Long - Kiên', 1],
  [3, 4, 'CD', 1, '8-11', 'Thúy - Phượng', 'Hà - Đạt', 1],
  [3, 4, 'CD', 2, '11-5', 'Toàn Nguyễn - Loan', 'Tùng - Dung', 1],
  [2, 5, 'AB', 1, '0-11', 'Hiếu - Sơn', 'Thắng - Long', 3],
  [2, 5, 'AB', 2, '11-7', 'Hải - Tú', 'Kỳ - Hiếu', 3],
  [2, 5, 'CD', 1, '11-8', 'Tới - Việt', 'Hoàng Thái - Ngọc Thắng', 3],
  [2, 5, 'CD', 2, '11-3', 'Thiên - Trung Anh', 'Văn Phương - Anh Bùi', 3],
  [1, 4, 'AB', 1, '11-4', 'Tuấn - Phúc', 'Tú - Quang', 4],
  [1, 4, 'AB', 2, '11-0', 'Nghị - Việt', 'Long - Kiên', 4],
  [1, 4, 'CD', 1, '9-11', 'Cường - Văn Anh', 'Hà - Đạt', 4],
  [1, 4, 'CD', 2, '11-3', 'Cường - Lan', 'Tùng - Dung', 4],
  [3, 5, 'AB', 1, '5-11', 'Toàn - Tiến', 'Long - Thắng', 5],
  [3, 5, 'AB', 2, '0-11', 'Tuấn - Vũ', 'Minh - Văn Thắng', 5],
  [3, 5, 'CD', 1, '9-11', 'Thúy - Phượng', 'Thái - Ngọc Thắng', 5],
  [3, 5, 'CD', 2, '11-3', 'Toàn Nguyễn - Loan', 'Nhàn - Thế Anh', 5],
  // Vòng 6: đội 2 – đội 4
  [2, 4, 'AB', 1, '11-5', 'Hiếu - Tú', 'Tú - Quang', 6],
  [2, 4, 'AB', 2, '11-4', 'Hải - Sơn', 'Long - Kiên', 6],
  [2, 4, 'CD', 1, '1-11', 'Tới - Việt', 'Hà - Tùng', 6],
  [2, 4, 'CD', 2, '11-1', 'Thiên - Trung Anh', 'Đạt - Dung', 6],
  // Vòng 7: đội 4 – đội 5 (hạng AB-2 chưa có tỉ số)
  [4, 5, 'AB', 1, '3-11', 'Tú - Long', 'Long - Thắng', 7],
  [4, 5, 'CD', 1, '11-7', 'Tùng - Hà', 'Thái - Ngọc Thắng', 7],
  [4, 5, 'CD', 2, '11-0', 'Đạt - Hà', 'Thế Anh - Minh Phương', 7],
  // Vòng 8: đội 1 – đội 3 (hạng AB-1 chưa có tỉ số)
  [1, 3, 'AB', 2, '11-0', 'Việt - Nghị', 'Toàn - Vũ', 8],
  [1, 3, 'CD', 1, '11-3', 'Cường - Lan', 'Thúy - Phượng', 8],
  [1, 3, 'CD', 2, '3-11', 'Cường - Văn Anh', 'Toàn Nguyễn - Loan', 8],
  // Bổ sung: vòng 7 hạng AB-2 (đội 4 – đội 5), vòng 8 hạng AB-1 (đội 1 – đội 3)
  [4, 5, 'AB', 2, '0-11', 'Quang - Kiến', 'Hiếu - Kỳ', 7],
  [1, 3, 'AB', 1, '11-7', 'Tuấn - Phúc', 'Toàn - Tiến', 8],
];
const resultOf = (t1, t2, tier, seq) => {
  const r = RESULTS.find(([a, b, h, s]) => a === t1 && b === t2 && h === tier && s === seq);
  return r ? r[4] : '';
};
const PLAYOFF_SLOTS = [['11:00', PH_THIRD], ['11:30', PH_FINAL]];
const DEFAULT_MATCHES = [
  ...SCHEDULE.map(([time, court, tier, seq, t1, t2], i) => ({
    id: 'M' + pad(i + 1), time, court, phase: PH_RR, tier, seq, t1, t2,
    score: resultOf(t1, t2, tier, seq), status: resultOf(t1, t2, tier, seq) ? 'Đã có tỉ số' : 'Dự kiến'
  })),
  ...PLAYOFF_SLOTS.flatMap(([time, phase]) => ['AB', 'CD'].flatMap(tier => [1, 2].map(seq => ({
    id: (phase === PH_FINAL ? 'CK' : 'TB') + '-' + tier + seq, time,
    court: 16 + ((tier === 'CD' ? 2 : 0) + seq - 1),
    phase, tier, seq, t1: null, t2: null, score: '', status: 'Dự kiến'
  }))))
];

const MATCH_COLS = ['Mã trận', 'Vòng', 'Hạng', 'Đội 1', 'Đội 2', 'Tỉ số', 'Trạng thái'];
const PLAYER_COLS = ['Trình', 'Họ tên', 'Phòng ban', 'Đội', 'Vai trò'];

/* ------------------------------------------------------------------ *
 *  3. HÀM TIỆN ÍCH
 * ------------------------------------------------------------------ */
const readSaved = (key, fallback) => {
  try { const v = localStorage.getItem(key); return v ? JSON.parse(v) : fallback; } catch { return fallback; }
};
const writeSaved = (key, value) => { try { localStorage.setItem(key, JSON.stringify(value)); } catch { /* bỏ qua */ } };
const clearSaved = () => { try { Object.values(KEYS).forEach(k => localStorage.removeItem(k)); } catch { /* bỏ qua */ } };

const parseScore = s => {
  const r = String(s ?? '').match(/(\d+)\s*[-–:]\s*(\d+)/);
  return r ? [Number(r[1]), Number(r[2])] : null;
};
const toTeam = v => { const n = String(v ?? '').match(/\d+/); return n ? Number(n[0]) : null; };
const toCourt = v => { const n = String(v ?? '').match(/\d+/); return n ? Number(n[0]) : ''; };
const fmtTime = v => {
  if (typeof v === 'number' && v >= 0 && v < 1) {
    const mins = Math.round(v * 1440);
    return pad(Math.floor(mins / 60)) + ':' + pad(mins % 60);
  }
  return String(v ?? '').trim();
};
const statusClass = s =>
  /đang/i.test(s || '') ? 'status status--live' : /kết thúc|hoàn thành|xong|đã có/i.test(s || '') ? 'status status--done' : 'status';

const parseHang = v => {
  const m = String(v ?? '').toUpperCase().match(/(AB|CD)\s*-?\s*(\d)/);
  return m ? { tier: m[1], seq: Number(m[2]) } : { tier: '', seq: 0 };
};
const hangLabel = m => `${m.tier}-${m.seq}`;
const normalizeMatch = r => ({
  id: String(r['Mã trận'] || '').trim(),
  time: fmtTime(r['Giờ']),
  court: toCourt(r['Sân']),
  phase: String(r['Vòng'] || PH_RR).trim(),
  ...parseHang(r['Hạng']),
  t1: toTeam(r['Đội 1']),
  t2: toTeam(r['Đội 2']),
  score: String(r['Tỉ số'] ?? '').trim(),
  status: String(r['Trạng thái'] || 'Dự kiến').trim()
});
const normalizePlayer = r => ({
  lvl: String(r['Trình'] || '').replace(/trình/i, '').trim().toUpperCase().slice(0, 1),
  name: String(r['Họ tên'] || '').trim(),
  dept: String(r['Phòng ban'] || '').trim(),
  team: toTeam(r['Đội']),
  captain: /đội trưởng/i.test(String(r['Vai trò'] || ''))
});

/* Xếp hạng vòng tròn: điểm đối đầu 3/1/0, rồi hiệu số trận, hiệu số điểm, đối đầu trực tiếp */
function computeStandings(rr) {
  const row = Object.fromEntries(TEAMS.map(t => [t, { team: t, pts: 0, played: 0, gw: 0, gl: 0, pf: 0, pa: 0 }]));
  const enc = {};
  rr.forEach(m => {
    const s = parseScore(m.score);
    if (!s || !m.t1 || !m.t2) return;
    const a = m.t1, b = m.t2;
    row[a].pf += s[0]; row[a].pa += s[1]; row[b].pf += s[1]; row[b].pa += s[0];
    if (s[0] > s[1]) { row[a].gw++; row[b].gl++; } else { row[b].gw++; row[a].gl++; }
    const k = encKey(a, b);
    if (!enc[k]) enc[k] = { lo: Math.min(a, b), hi: Math.max(a, b), n: 0, wins: {} };
    enc[k].n++;
    const w = s[0] > s[1] ? a : b;
    enc[k].wins[w] = (enc[k].wins[w] || 0) + 1;
  });
  const encRes = {};
  Object.entries(enc).forEach(([k, e]) => {
    const complete = e.n === 4;
    const wl = e.wins[e.lo] || 0, wh = e.wins[e.hi] || 0;
    const pts = { [e.lo]: 0, [e.hi]: 0 };
    if (complete) {
      if (wl === wh) { pts[e.lo] = 1; pts[e.hi] = 1; }
      else if (wl > wh) pts[e.lo] = 3; else pts[e.hi] = 3;
      [e.lo, e.hi].forEach(t => { row[t].pts += pts[t]; row[t].played++; });
    }
    encRes[k] = { complete, pts, wl, wh, n: e.n, lo: e.lo, hi: e.hi };
  });
  const pointsIn = (a, b) => { const r = encRes[encKey(a, b)]; return r && r.complete ? r.pts[a] : 0; };
  const rows = Object.values(row).sort((x, y) =>
    (y.pts - x.pts) ||
    ((y.gw - y.gl) - (x.gw - x.gl)) ||
    ((y.pf - y.pa) - (x.pf - x.pa)) ||
    (pointsIn(y.team, x.team) - pointsIn(x.team, y.team)) ||
    (x.team - y.team));
  return { rows, encRes };
}

// Dựng dữ liệu hiển thị: điền đội vào trận tranh hạng khi vòng tròn xong
function buildView(matches) {
  const rr = matches.filter(m => m.phase === PH_RR);
  const complete = rr.length > 0 && rr.every(m => parseScore(m.score));
  const { rows, encRes } = computeStandings(rr);
  const rank = i => (complete && rows[i] ? rows[i].team : null);
  const seats = { [PH_FINAL]: [0, 1], [PH_THIRD]: [2, 3] };
  const view = matches.map(m => {
    if (!seats[m.phase]) return m;
    const [i, j] = seats[m.phase];
    return { ...m, t1: m.t1 ?? rank(i), t2: m.t2 ?? rank(j), seat1: i + 1, seat2: j + 1 };
  });
  return { view, rows, encRes, complete };
}

const fmtDiff = n => (n > 0 ? '+' : '') + n;

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
 *  5. GIAO DIỆN
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

function Hero() {
  const cd = useCountdown(EVENT.startISO);
  return (
    <section className="hero" id="top">
      <div className="container hero-grid">
        <div className="hero-copy">
          <p className="hero-org">MSB · Strategy &amp; Innovation</p>
          <h1><span>Pickleball</span><span className="h1-band">SI Open 2026</span></h1>
          <p className="hero-tag">Bứt phá giới hạn – Gắn kết thành công</p>
          <p className="hero-lead">5 đội, 43 vận động viên và 6 sân thi đấu song song. Cả Khối SI cùng đến cổ vũ cho đồng đội.</p>
          <div className="hero-actions">
            <a className="btn btn-accent" href="#lich">Tra cứu trận của bạn <ArrowUpRight size={18} aria-hidden="true" /></a>
            <a className="btn btn-ghost" href="#doi-thi">Danh sách VĐV</a>
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
          <div><span className="fact-icon"><CalendarDays size={20} aria-hidden="true" /></span><div><dt>Ngày thi đấu</dt><dd>{EVENT.dateLabel}<small>{EVENT.venue}</small></dd></div></div>
          <div><span className="fact-icon"><Users size={20} aria-hidden="true" /></span><div><dt>Thể thức</dt><dd>5 đội · 43 VĐV<small>40 trận vòng tròn</small></dd></div></div>
          <div><span className="fact-icon"><MapPin size={20} aria-hidden="true" /></span><div><dt>Địa điểm</dt><dd><a href={MAPS_URL} target="_blank" rel="noopener noreferrer">{EVENT.venue}</a><small>{EVENT.address}</small></dd></div></div>
          <div><span className="fact-icon"><Trophy size={20} aria-hidden="true" /></span><div><dt>Tổng giải thưởng</dt><dd>{EVENT.prizeTotal}</dd></div></div>
        </dl>
      </div>
    </section>
  );
}

function Players({ players }) {
  const count = l => players.filter(p => p.lvl === l).length;
  return (
    <section className="section" id="doi-thi">
      <div className="container">
        <SectionHead title="Vận động viên và đội thi đấu"
          lead="43 vận động viên chia theo 4 trình A–D. Đội 1–4 mỗi đội 8 người, Đội 5 là đội khách mời 11 người. Chữ cái bên cạnh tên là trình."
          aside={(
            <div className="progress">
              <p><b>{players.length}</b><span>/43 VĐV trong danh sách</span></p>
              <div className="progress-track" role="progressbar" aria-valuemin={0} aria-valuemax={43} aria-valuenow={players.length}>
                <i style={{ width: Math.min(100, Math.round((players.length / 43) * 100)) + '%' }} />
              </div>
              <small>{LEVELS.map(l => `${l} ${count(l)}`).join(' · ')}</small>
            </div>
          )} />

        <div className="levels">
          {LEVELS.map(l => {
            const list = players.filter(p => p.lvl === l);
            return (
              <article className="level" key={l}>
                <header className="level-head">
                  <span className="level-letter" aria-hidden="true">{l}</span>
                  <h3>Trình {l}</h3>
                  <span className="level-count">{list.length} VĐV</span>
                </header>
                <ol className="level-list">
                  {list.map((p, i) => (
                    <li className={'is-filled' + (isGuest(p.team) ? ' is-guest' : '')} key={p.name + i}>
                      <span className="pname">{p.name}{p.captain && <abbr className="dt" title="Đội trưởng">ĐT</abbr>}</span>
                      <span className="pdept" title={p.dept}>{teamName(p.team)}{p.dept ? ' · ' + p.dept : ''}</span>
                    </li>
                  ))}
                </ol>
              </article>
            );
          })}
        </div>

        <div className="teams-strip">
          <div>
            <h3 className="subhead">Chia đội</h3>
            <p className="subnote">Đội 1–4 là đội nội bộ, mỗi đội 8 người. Đội 5 là đội khách mời 11 người.</p>
          </div>
          <div className="team-cards">
            {TEAMS.map(t => {
              const all = players.filter(p => p.team === t)
                .sort((a, b) => LEVELS.indexOf(a.lvl) - LEVELS.indexOf(b.lvl));
              const guest = isGuest(t);
              return (
                <article className={'tcard' + (guest ? ' is-guest' : '')} key={t}>
                  <header>
                    <span className="team-no" aria-hidden="true">{t}</span>
                    <h4>{teamName(t)}</h4>
                    {guest && <span className="tguest">Khách mời</span>}
                    <span className="tcount">{all.length}/{guest ? EVENT.guestSize : EVENT.teamSize}</span>
                  </header>
                  <ol className="tlist">
                    {all.map((m, k) => (
                      <li key={m.name + k}>
                        <span className="tname">{m.name}</span>
                        <span className="tmeta">
                          {m.captain && <abbr className="dt" title="Đội trưởng">ĐT</abbr>}
                          <i className="lv" title={'Trình ' + m.lvl}>{m.lvl}</i>
                        </span>
                      </li>
                    ))}
                  </ol>
                </article>
              );
            })}
          </div>
        </div>
      </div>
    </section>
  );
}

function Register() {
  return (
    <section className="register" id="dang-ky">
      <div className="container register-grid">
        <div>
          <h2>Danh sách đăng ký</h2>
          <p>Danh sách đăng ký hiện đã đủ số vận động viên. Vui lòng liên hệ Ban Tổ chức nếu có thay đổi.</p>
          <div className="register-actions">
            <a className="btn btn-line" href="#doi-thi">Xem danh sách VĐV</a>
            <a className="btn btn-line" href="#lich">Xem lịch thi đấu</a>
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

function Timeline() {
  return (
    <section className="section section--paper" id="lich-trinh">
      <div className="container">
        <SectionHead title="Chương trình" lead="Các bước chính của giải. Ban Tổ chức sẽ thông báo chi tiết trước ngày thi đấu." />
        <ol className="timeline">
          {TIMELINE.map(([time, tag, title, text]) => (
            <li className="tl-item" key={time + tag}>
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

const resultInfo = (t1, t2, tier, seq) => RESULTS.find(([a, b, h, s]) => a === t1 && b === t2 && h === tier && s === seq);
const vongOf = m => { const r = resultInfo(m.t1, m.t2, m.tier, m.seq); return r ? r[7] : 99; };

function ScoreTables({ view }) {
  const rr = view.filter(m => m.phase === PH_RR);
  return (
    <section className="section" id="lich">
      <div className="container">
        <SectionHead title="Bảng tỉ số vòng tròn"
          lead="Mỗi cặp đội có 4 hạng: AB-1, AB-2, CD-1, CD-2. Mỗi hạng một bảng, ghi rõ vòng, đội, tên VĐV và tỉ số." />
        <div className="score-grid">
          {[['AB', 1], ['AB', 2], ['CD', 1], ['CD', 2]].map(([tier, seq]) => {
            const list = rr.filter(m => m.tier === tier && m.seq === seq)
              .sort((a, b) => (vongOf(a) - vongOf(b)) || a.t1 - b.t1 || a.t2 - b.t2);
            return (
              <article className="score-card" key={tier + seq}>
                <h3>Hạng {tier}-{seq}</h3>
                <table className="table score-table">
                  <caption className="sr-only">Tỉ số hạng {tier}-{seq}</caption>
                  <thead>
                    <tr><th scope="col">Vòng</th><th scope="col">Đội 1</th><th scope="col" className="num">Tỉ số</th><th scope="col">Đội 2</th></tr>
                  </thead>
                  <tbody>
                    {list.map(m => {
                      const info = resultInfo(m.t1, m.t2, tier, seq);
                      const sc = parseScore(m.score);
                      return (
                        <tr key={m.id}>
                          <td className="rank">{info ? info[7] : '–'}</td>
                          <td><span className="t-name">{teamName(m.t1)}</span><span className="t-sub">{info ? info[5] : ''}</span></td>
                          <td className="num"><b>{sc ? `${sc[0]} – ${sc[1]}` : '–'}</b></td>
                          <td><span className="t-name">{teamName(m.t2)}</span><span className="t-sub">{info ? info[6] : ''}</span></td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </article>
            );
          })}
        </div>
      </div>
    </section>
  );
}

function Standings({ rows, encRes, view }) {
  const anyPlayed = view.some(m => m.phase === PH_RR && parseScore(m.score));
  return (
    <section className="section section--paper" id="bang">
      <div className="container">
        <SectionHead title="Bảng xếp hạng vòng tròn"
          lead="Mỗi cuộc đối đầu có 4 trận: thắng 3 điểm, hoà 2–2 mỗi đội 1 điểm, thua 0 điểm. Khi bằng điểm, xét hiệu số trận, hiệu số điểm, rồi đối đầu trực tiếp."
          aside={!anyPlayed && <p className="pill">Chưa có kết quả</p>} />
        <div className="two-col">
          <div>
            <div className="table-wrap">
              <table className="table">
                <caption className="sr-only">Bảng xếp hạng các đội</caption>
                <thead>
                  <tr>
                    <th scope="col">Hạng</th><th scope="col">Đội</th>
                    <th scope="col" className="num">Điểm</th>
                    <th scope="col" className="num">Đối đầu</th>
                    <th scope="col" className="num">Trận T–B</th>
                    <th scope="col" className="num">Điểm ghi–thủng</th>
                    <th scope="col" className="num">Hiệu số</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r, i) => (
                    <tr key={r.team} className={i < 2 ? 'row-top' : i < 4 ? 'row-low' : 'row-out'}>
                      <td className="rank">{anyPlayed ? i + 1 : '–'}</td>
                      <th scope="row"><span className="t-name">{teamName(r.team)}</span>{isGuest(r.team) && <span className="t-sub">Khách mời</span>}</th>
                      <td className="num"><b>{r.pts}</b></td>
                      <td className="num">{r.played}<span className="of">/4</span></td>
                      <td className="num">{r.gw}–{r.gl}</td>
                      <td className="num">{r.pf}–{r.pa}</td>
                      <td className="num">{fmtDiff(r.pf - r.pa)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <p className="legend"><i className="dot dot--top" /> Hạng 1–2 đấu chung kết <i className="dot dot--low" /> Hạng 3–4 tranh giải Ba <i className="dot dot--out" /> Hạng 5 dừng ở vòng tròn</p>
          </div>

          <div>
            <h3 className="subhead">Bảng đối đầu</h3>
            <p className="subnote">Đọc theo hàng: số trận thắng–thua của đội ở hàng khi gặp đội ở cột (trên tổng 4 trận).</p>
            <div className="table-wrap">
              <table className="table cross">
                <caption className="sr-only">Bảng đối đầu giữa các đội</caption>
                <thead><tr><td />{TEAMS.map(t => <th scope="col" key={t}>{teamName(t)}</th>)}</tr></thead>
                <tbody>
                  {TEAMS.map(a => (
                    <tr key={a}>
                      <th scope="row">{teamName(a)}</th>
                      {TEAMS.map(b => {
                        if (a === b) return <td key={b} className="self" aria-hidden="true" />;
                        const e = encRes[encKey(a, b)];
                        if (!e || !e.n) return <td key={b} className="cell-empty">–</td>;
                        const wa = a === e.lo ? e.wl : e.wh, wb = a === e.lo ? e.wh : e.wl;
                        if (!e.complete) return <td key={b} className="cell-plan"><b>{wa}–{wb}</b><span>{e.n}/4 trận</span></td>;
                        const pa = e.pts[a];
                        return (
                          <td key={b} className={'cell-score ' + (pa === 3 ? 'is-win' : pa === 1 ? 'cell-part' : 'is-loss')}>
                            <b>{wa}–{wb}</b><span>{pa} điểm</span>
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

function BracketMatch({ title, a, b, note }) {
  return (
    <div className="bk-match">
      <p className="bk-title">{title}</p>
      {[a, b].map((s, i) => (
        <div key={i} className={'bk-row' + (s.win ? ' is-win' : '') + (s.pending ? ' is-pending' : '')}>
          <span>{s.name}</span><b>{s.score ?? '–'}</b>
        </div>
      ))}
      {note && <p className="bk-note">{note}</p>}
    </div>
  );
}

function Bracket({ view, rows }) {
  const tally = phase => {
    const ms = view.filter(m => m.phase === phase);
    const first = ms[0] || {};
    let w1 = 0, w2 = 0, done = 0;
    ms.forEach(m => {
      const s = parseScore(m.score);
      if (!s) return;
      done++;
      if (s[0] > s[1]) w1++; else if (s[1] > s[0]) w2++;
    });
    const finished = ms.length > 0 && done === ms.length;
    const nameOf = (t, seat) => (t ? teamName(t) : `Hạng ${seat ?? ''}`);
    return {
      a: { name: nameOf(first.t1, first.seat1), score: done ? w1 : undefined, win: finished && w1 > w2, pending: !first.t1 },
      b: { name: nameOf(first.t2, first.seat2), score: done ? w2 : undefined, win: finished && w2 > w1, pending: !first.t2 },
      finished,
      champion: finished && w1 !== w2 ? (w1 > w2 ? first.t1 : first.t2) : null,
      tie: finished && w1 === w2
    };
  };
  const fin = tally(PH_FINAL);
  const third = tally(PH_THIRD);
  const champion = fin.champion;
  return (
    <section className="section" id="nhanh">
      <div className="container">
        <SectionHead title="Con đường tới ngôi vô địch"
          lead="Sau vòng tròn, hạng 3–4 đấu tranh giải Ba; hạng 1–2 đấu chung kết. Mỗi trận gồm 4 hạng, tên đội tự điền theo bảng xếp hạng." />
        <div className="bracket">
          <BracketMatch title="Chung kết · 4 hạng" a={fin.a} b={fin.b}
            note={fin.tie ? 'Tỉ số hoà 2–2: BTC quyết định thêm.' : null} />
          <span className="bk-link" aria-hidden="true" />
          <div className="bk-champ">
            <Trophy size={34} aria-hidden="true" />
            <strong>{champion ? teamName(champion) : 'Nhà vô địch'}</strong>
            <span>{champion ? 'SI Open 2026' : 'Chờ trận chung kết'}</span>
          </div>
          <div className="bk-third">
            <BracketMatch title="Tranh giải Ba · 4 hạng" a={third.a} b={third.b}
              note={third.tie ? 'Tỉ số hoà 2–2: BTC quyết định thêm.' : null} />
          </div>
        </div>
        {rows.length === 0 && null}
      </div>
    </section>
  );
}

function Rules() {
  return (
    <section className="section section--paper" id="the-le">
      <div className="container rules-layout">
        <SectionHead title="Thể lệ thi đấu" lead="Điều lệ và quy chế thi đấu của giải. Bấm từng mục để xem chi tiết." />
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

function Cheer() {
  return (
    <section className="cheer" id="co-vu">
      <div className="container cheer-grid">
        <div>
          <h2>Đến thật đông, cổ vũ thật to!</h2>
          <p>Giải của cả Khối SI. Ai thi đấu thì ra sân cùng đồng đội, ai chưa thi đấu thì đến cổ vũ. Một tràng pháo tay cũng là cả một trận thắng.</p>
        </div>
        <div className="cheer-tiles">
          <div className="cheer-tile"><Users size={28} aria-hidden="true" /><div><h3>Thi đấu</h3><p>Xem trận của đội và của bạn trong mục lịch thi đấu.</p></div></div>
          <div className="cheer-tile"><Megaphone size={28} aria-hidden="true" /><div><h3>Cổ vũ</h3><p>Đến {EVENT.venue}, {EVENT.dateLabel} để tiếp lửa cho đồng đội.</p></div></div>
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
        <p>{EVENT.dateLabel}<br />
          <a href={MAPS_URL} target="_blank" rel="noopener noreferrer">{EVENT.venue}, {EVENT.address}</a></p>
        <div className="closing-actions">
          <a className="btn btn-accent" href="#lich">Tìm trận của bạn <ArrowUpRight size={18} aria-hidden="true" /></a>
          <a className="btn btn-ghost" href="#doi-thi">Xem danh sách VĐV</a>
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
      const ms = sheet('Lịch thi đấu').map(normalizeMatch).filter(x => x.id);
      const ps = sheet('Vận động viên').map(normalizePlayer).filter(p => p.name && p.team);
      if (!ms.length && !ps.length) {
        setFile(null);
        setNotice('Không tìm thấy dữ liệu. File cần có sheet “Lịch thi đấu” hoặc “Vận động viên”.');
        return;
      }
      setFile({ matches: ms, players: ps });
      setNotice(`Đã đọc ${ms.length} trận và ${ps.length} VĐV. Bấm “Áp dụng” để xem trên trình duyệt này.`);
    } catch {
      setFile(null);
      setNotice('Không đọc được file. Hãy kiểm tra lại định dạng .xlsx hoặc .xls.');
    }
  };

  const download = () => {
    const wb = XLSX.utils.book_new();
    const mRows = matches.map(m => ({
      'Mã trận': m.id, 'Vòng': m.phase, 'Hạng': hangLabel(m),
      'Đội 1': m.t1 ?? '', 'Đội 2': m.t2 ?? '', 'Tỉ số': m.score, 'Trạng thái': m.status
    }));
    const pRows = players.map(p => ({
      'Trình': p.lvl, 'Họ tên': p.name, 'Phòng ban': p.dept, 'Đội': p.team, 'Vai trò': p.captain ? 'Đội trưởng' : ''
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(mRows, { header: MATCH_COLS }), 'Lịch thi đấu');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(pRows, { header: PLAYER_COLS }), 'Vận động viên');
    XLSX.writeFile(wb, 'SI_Open_2026_du_lieu.xlsx');
  };

  return (
    <section className="admin" id="btc">
      <div className="container">
        <div className="admin-head">
          <h2>Khu vực Ban Tổ chức</h2>
          <button type="button" className="icon-btn" onClick={onClose} aria-label="Đóng khu vực Ban Tổ chức"><X size={20} /></button>
        </div>
        <p className="admin-lead">Tải file mẫu về (đang có sẵn lịch và danh sách), điền tiếp rồi tải lên. Sheet <b>“Lịch thi đấu”</b>: Tỉ số (dạng <b>11-8</b>, đặt ô là Text), Trạng thái.
          Cột <b>Đội 1</b>, <b>Đội 2</b> là số đội (1–5). Sheet <b>“Vận động viên”</b>: Trình (A–D), Họ tên, Phòng ban, Đội (1–5), Vai trò (ghi <b>Đội trưởng</b> cho đội trưởng).</p>

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
  const { view, rows, encRes } = useMemo(() => buildView(matches), [matches]);

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
        <Players players={players} />
        <Register />
        <Timeline />
        <ScoreTables view={view} />
        <Standings rows={rows} encRes={encRes} view={view} />
        <Bracket view={view} rows={rows} />
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
