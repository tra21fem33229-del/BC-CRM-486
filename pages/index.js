import { useEffect, useState, useMemo } from 'react';
import Link from 'next/link';

function rankBadge(i) {
  if (i === 0) return <span className="medal">🥇</span>;
  if (i === 1) return <span className="medal">🥈</span>;
  if (i === 2) return <span className="medal">🥉</span>;
  return i + 1;
}

// Thư viện xlsx (~500KB) chỉ nạp khi người dùng thực sự bấm nút xuất, tránh làm nặng trang công khai.
let xlsxModulePromise = null;
function loadXLSX() {
  if (!xlsxModulePromise) xlsxModulePromise = import('xlsx');
  return xlsxModulePromise;
}

async function exportXlsx(rows, kind, monthLabel, cdsOnly) {
  if (!rows.length) return;
  const XLSX = await loadXLSX();
  const isPhong = kind === 'phong';
  const suf = cdsOnly ? 'CDS' : '';
  const val = (r, field) => r[field + suf] ?? 0;
  const diemVal = (r) => (cdsOnly ? r.diemCDS : r.diem);

  const header = isPhong
    ? ['Hạng', 'Phòng / PGD', 'Số RM', 'Lead giao', 'Lead/Opp có tương tác', 'Lead → Opp', 'Opp thành công', 'Điểm thi đua']
    : ['Hạng', 'Cán bộ (RM)', 'Phòng', 'Lead giao', 'Lead/Opp có tương tác', 'Lead → Opp', 'Opp thành công', 'Điểm thi đua'];

  const body = rows.map((r, i) => [
    diemVal(r) === null || diemVal(r) === undefined ? '' : i + 1,
    isPhong ? r.label || r.key : r.key,
    isPhong ? r.soRM || '' : r.phongLabel || r.phong || '',
    val(r, 'leadGiao'),
    val(r, 'leadTuongTac') + val(r, 'oppTuongTac'),
    val(r, 'leadChuyenDoi'),
    val(r, 'oppThanhCong'),
    diemVal(r) === null || diemVal(r) === undefined ? '' : diemVal(r),
  ]);

  const tieuDe =
    (isPhong ? 'XẾP HẠNG THEO PHÒNG / PGD' : 'XẾP HẠNG THEO CÁN BỘ (RM)') + (cdsOnly ? ' — CHỈ SÁNG KIẾN CĐS' : '');
  const aoa = [
    ['VIETINBANK — CHI NHÁNH NGŨ HÀNH SƠN'],
    ['Bảng điểm thi đua CRM 1.0 Transformation 2026 — ' + tieuDe],
    ['Kỳ xét thưởng: ' + (monthLabel || '—')],
    [],
    header,
    ...body,
  ];

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = header.map((h, i) => ({ wch: i === 1 ? 34 : Math.max(12, h.length + 3) }));
  ws['!merges'] = [0, 1, 2].map((r) => ({ s: { r, c: 0 }, e: { r, c: header.length - 1 } }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, isPhong ? 'Phong-PGD' : 'Can bo RM');

  const stamp = (monthLabel || 'du-lieu').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^0-9A-Za-z]+/g, '-');
  XLSX.writeFile(wb, `Thi-dua-CRM_${isPhong ? 'Phong' : 'RM'}_${cdsOnly ? 'CDS_' : ''}${stamp}.xlsx`);
}

function DataTable({ rows, kind, cdsOnly }) {
  if (!rows.length) {
    return (
      <div className="panel">
        <div className="empty-state">
          <div className="big">—</div>
          Chưa có dữ liệu phù hợp.
        </div>
      </div>
    );
  }
  const nameHeader = kind === 'phong' ? 'Phòng / PGD' : 'Cán bộ (RM)';
  const suf = cdsOnly ? 'CDS' : '';
  const val = (r, field) => r[field + suf] ?? 0;
  const diemVal = (r) => (cdsOnly ? r.diemCDS : r.diem);
  return (
    <div className="panel">
      <table>
        <thead>
          <tr>
            <th style={{ width: 34 }}>#</th>
            <th>{nameHeader}</th>
            {kind === 'rm' && <th>Phòng</th>}
            {kind === 'phong' && <th>Số RM</th>}
            <th>Lead giao</th>
            <th>Lead/Opp <br />có tương tác</th>
            <th>Lead → Opp</th>
            <th>Opp <br />thành công</th>
            <th>Điểm thi đua{cdsOnly ? ' (CĐS)' : ''}</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={r.key}>
              <td className="rank">{diemVal(r) === null || diemVal(r) === undefined ? '—' : rankBadge(i)}</td>
              <td className="name-cell">{kind === 'phong' ? r.label || r.key : r.key}</td>
              {kind === 'rm' && <td>{r.phongLabel || r.phong || '—'}</td>}
              {kind === 'phong' && <td className="mono">{r.soRM || '—'}</td>}
              <td className="mono">{val(r, 'leadGiao').toLocaleString('vi-VN')}</td>
              <td className="mono">{(val(r, 'leadTuongTac') + val(r, 'oppTuongTac')).toLocaleString('vi-VN')}</td>
              <td className="mono">{val(r, 'leadChuyenDoi')}</td>
              <td className="mono">{val(r, 'oppThanhCong')}</td>
              <td className="diem mono">{diemVal(r) === null || diemVal(r) === undefined ? '—' : diemVal(r)}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export default function Home() {
  const [months, setMonths] = useState([]);
  const [selected, setSelected] = useState('');
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [query, setQuery] = useState('');
  const [cdsOnly, setCdsOnly] = useState(false);

  useEffect(() => {
    fetch('/api/months')
      .then((r) => r.json())
      .then((res) => {
        const list = res.months || [];
        setMonths(list);
        if (list.length) setSelected(list[list.length - 1].key);
        setLoading(false);
      })
      .catch(() => setLoading(false));
  }, []);

  useEffect(() => {
    if (!selected) return;
    fetch(`/api/data/${encodeURIComponent(selected)}`)
      .then((r) => r.json())
      .then((res) => setData(res.data))
      .catch(() => setData(null));
  }, [selected]);

  const filtered = useMemo(() => {
    if (!data) return { phong: [], rm: [] };
    const q = query.trim().toLowerCase();
    const filt = (arr) =>
      q
        ? arr.filter(
            (r) =>
              r.key.toLowerCase().includes(q) ||
              (r.label || '').toLowerCase().includes(q) ||
              (r.phong || '').toLowerCase().includes(q) ||
              (r.phongLabel || '').toLowerCase().includes(q)
          )
        : arr;
    // Khi bật "Sáng kiến CĐS", sắp xếp lại theo diemCDS (dữ liệu gốc từ DB luôn sắp theo điểm
    // đầy đủ) — cùng quy tắc null-cuối-danh-sách như scoreRM/scorePhong.
    const sortByDiemCDS = (arr) =>
      [...arr].sort((a, b) => {
        const da = a.diemCDS,
          db = b.diemCDS;
        if (da === null || da === undefined) return db === null || db === undefined ? 0 : 1;
        if (db === null || db === undefined) return -1;
        return db - da;
      });
    let phong = filt(data.phong || []);
    let rm = filt(data.rm || []);
    if (cdsOnly) {
      phong = sortByDiemCDS(phong);
      rm = sortByDiemCDS(rm);
    }
    return { phong, rm };
  }, [data, query, cdsOnly]);

  const monthLabel =
    selected === '__all__'
      ? 'Lũy kế tất cả các kỳ'
      : months.find((m) => m.key === selected)?.label || selected;

  const s = data?.summary || {};
  const sVal = (field) => (cdsOnly ? s[field + 'CDS'] || 0 : s[field] || 0);
  const tyLeChung = sVal('leadGiao') ? ((100 * sVal('leadTuongTac')) / sVal('leadGiao')).toFixed(1) : '0.0';

  return (
    <>
      <div className="topbar">
        <div className="topbar-inner">
          <div>
            <div className="eyebrow">VietinBank · Chi nhánh Ngũ Hành Sơn</div>
            <div className="title">CRM1.0 Transformation 2026 — Đổi hành vi, tăng hiệu quả</div>
            <div className="subtitle">
              Bảng điểm thi đua triển khai CRM 1.0 — cập nhật theo từng kỳ tháng, tính theo
              Phòng/PGD và cán bộ RM.
            </div>
          </div>
          <div className="nav">
            <Link href="/" className="active">
              Bảng xếp hạng
            </Link>
            <Link href="/canh-bao">Cảnh báo</Link>
            <Link href="/admin">Quản trị</Link>
          </div>
        </div>
      </div>

      <div className="wrap">
        <div className="toolbar">
          <div>
            <span className="field-label">Kỳ xét thưởng</span>
            <select value={selected} onChange={(e) => setSelected(e.target.value)}>
              {!months.length && <option value="">Chưa có dữ liệu</option>}
              {months
                .slice()
                .reverse()
                .map((m) => (
                  <option key={m.key} value={m.key}>
                    {m.label}
                  </option>
                ))}
              {months.length > 0 && <option value="__all__">Lũy kế tất cả các kỳ</option>}
            </select>
          </div>
          <div className="topright" style={{ marginLeft: 'auto', display: 'flex', gap: 20, alignItems: 'flex-end' }}>
            <div>
              <span className="field-label">Nguồn dữ liệu</span>
              <button
                type="button"
                className={`btn ${cdsOnly ? '' : 'secondary'} cds-toggle`}
                aria-pressed={cdsOnly}
                title="Chỉ tính Lead/Opp thuộc Nhóm nguồn &quot;Kênh sáng kiến CĐS KHBL&quot; và &quot;Kênh sáng kiến CĐS KHDN&quot;"
                onClick={() => setCdsOnly((v) => !v)}
              >
                {cdsOnly ? '✓ Sáng kiến CĐS' : 'Sáng kiến CĐS'}
              </button>
            </div>
            <div className="search-box">
              <span className="field-label">Tìm phòng / cán bộ</span>
              <input
                type="text"
                placeholder="Nhập tên phòng hoặc mã RM..."
                value={query}
                onChange={(e) => setQuery(e.target.value)}
              />
            </div>
          </div>
        </div>
        {cdsOnly && (
          <div className="msg success" style={{ marginTop: -6, marginBottom: 16 }}>
            Đang chỉ tính Lead/Opp thuộc nhóm nguồn "Kênh sáng kiến CĐS KHBL" và "Kênh sáng kiến CĐS KHDN". Bấm lại nút để xem toàn bộ số liệu.
          </div>
        )}

        {loading ? (
          <div className="panel">
            <div className="empty-state">Đang tải dữ liệu...</div>
          </div>
        ) : !selected || !data ? (
          <div className="panel">
            <div className="empty-state">
              <div className="big">📊</div>
              Chưa có dữ liệu kỳ nào được tải lên.
              <br />
              Vui lòng liên hệ Phòng Kế hoạch Tổng hợp để cập nhật.
            </div>
          </div>
        ) : (
          <>
            <div className="stat-grid">
              <div className="stat-card">
                <div className="num">{sVal('leadGiao').toLocaleString('vi-VN')}</div>
                <div className="lbl">Lead đã phân giao</div>
              </div>
              <div className="stat-card">
                <div className="num">{sVal('leadTuongTac').toLocaleString('vi-VN')}</div>
                <div className="lbl">Lead có tương tác</div>
              </div>
              <div className="stat-card accent">
                <div className="num">{tyLeChung}%</div>
                <div className="lbl">Tỷ lệ tiếp cận chung</div>
              </div>
              <div className="stat-card">
                <div className="num">{sVal('leadChuyenDoi').toLocaleString('vi-VN')}</div>
                <div className="lbl">Lead chuyển đổi → Opp</div>
              </div>
              <div className="stat-card">
                <div className="num">{sVal('oppThanhCong').toLocaleString('vi-VN')}</div>
                <div className="lbl">Opp thành công</div>
              </div>
              <div className="stat-card">
                <div className="num">{(s.tongRM || 0).toLocaleString('vi-VN')}</div>
                <div className="lbl">Tổng số RM</div>
              </div>
            </div>

            <div className="section-title">
              <h2>Xếp hạng theo Phòng / PGD{cdsOnly ? ' — Sáng kiến CĐS' : ''}</h2>
              <span className="count-pill">{filtered.phong.length}</span>
              <button
                className="btn secondary export-btn"
                onClick={() => exportXlsx(filtered.phong, 'phong', monthLabel, cdsOnly)}
                disabled={!filtered.phong.length}
              >
                ⬇ Xuất Excel
              </button>
            </div>
            <DataTable rows={filtered.phong} kind="phong" cdsOnly={cdsOnly} />

            <div className="section-title">
              <h2>Xếp hạng theo Cán bộ (RM){cdsOnly ? ' — Sáng kiến CĐS' : ''}</h2>
              <span className="count-pill">{filtered.rm.length}</span>
              <button
                className="btn secondary export-btn"
                onClick={() => exportXlsx(filtered.rm, 'rm', monthLabel, cdsOnly)}
                disabled={!filtered.rm.length}
              >
                ⬇ Xuất Excel
              </button>
            </div>
            <DataTable rows={filtered.rm} kind="rm" cdsOnly={cdsOnly} />
          </>
        )}
      </div>

      <footer>VietinBank Chi nhánh Ngũ Hành Sơn · Dữ liệu phục vụ chương trình thi đua CRM1.0 Transformation 2026</footer>
    </>
  );
}
