"""
BFD Diagram HTML Generator
Produces a self-contained HTML visual of the Block Flow Diagram.
Columns: IS-xx (left) → Block (centre) → OS-xx/PS-xx (right)
Waste stream color coding matches the Excel standard.
"""

WASTE_COLORS = {
    "gas_waste"    : ("#dbeafe", "#1e40af", "Gas Waste"),
    "organic_waste": ("#fef9c3", "#854d0e", "Organic Waste"),
    "aqueous_waste": ("#fce7f3", "#9d174d", "Aqueous Waste"),
    "solid_waste"  : ("#ffedd5", "#9a3412", "Solid Waste"),
}


def generate_diagram_html(data):
    project      = data.get("project", {})
    calc         = data.get("calc", {})
    product_name = project.get("product_name", "Product")
    batch_size   = project.get("batch_size", 0)
    reference    = project.get("reference", "")
    ops          = calc.get("operations", data.get("operations", []))
    yld          = calc.get("yield", {})
    calc_comps   = calc.get("components", data.get("components", []))

    # ── Yield banner ───────────────────────────────────────────────
    yield_html = ""
    if yld:
        pct    = yld.get("molar_yield_pct", "—")
        color  = ("#065f46" if float(pct or 0) >= 85
                  else "#92400e" if float(pct or 0) >= 70
                  else "#991b1b")
        yield_html = f"""
        <div class="yield-banner">
          <div class="y-stat"><div class="y-val">{yld.get('sm_input_kg','—')}</div><div class="y-lbl">SM Input (kg)</div></div>
          <div class="y-arrow">→</div>
          <div class="y-stat"><div class="y-val">{yld.get('product_output_kg','—')}</div><div class="y-lbl">Product Output (kg)</div></div>
          <div class="y-arrow">→</div>
          <div class="y-stat"><div class="y-val" style="color:{color}">{pct}%</div><div class="y-lbl">Molar Yield</div></div>
        </div>"""

    # ── RM table rows ──────────────────────────────────────────────
    rm_rows = ""
    for c in calc_comps:
        role       = c.get("role", "reagent").lower()
        row_cls    = {"sm":"row-sm","starting_material":"row-sm",
                      "product":"row-product","solvent":"row-solvent",
                      "byproduct":"row-byproduct"}.get(role, "")
        rm_rows += f"""<tr class="{row_cls}">
          <td>{c.get('name','')}</td><td>{c.get('mw','')}</td>
          <td>{c.get('molar_ratio','')}</td><td>{c.get('moles','')}</td>
          <td>{c.get('purity',100)}%</td>
          <td><b>{c.get('mass_kg','')}</b></td>
          <td>{c.get('density','')}</td><td>{c.get('volume_kl','')}</td>
        </tr>"""

    # ── Operation blocks ───────────────────────────────────────────
    blocks_html = "".join(_render_block(op) for op in ops)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>BFD — {product_name}</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');
:root{{
  --navy:#0f2744;--blue:#1e5aa0;--accent:#2e9afe;
  --block:#b4c6e7;--block-txt:#0f2744;
  --border:#8eaac8;--text:#1a2332;--muted:#607080;
  --ps-bg:#e2efda;--ps-border:#5cb85c;
  --comp-bg:#f5f9ff;
}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'IBM Plex Sans',sans-serif;background:#e8f0f8;color:var(--text);min-height:100vh}}
.page-hdr{{background:linear-gradient(135deg,var(--navy),var(--blue));color:white;
  padding:20px 36px;display:flex;justify-content:space-between;align-items:center;
  box-shadow:0 4px 16px rgba(0,0,0,.25)}}
.page-hdr h1{{font-size:20px;font-weight:700;letter-spacing:-.3px}}
.page-hdr .meta{{font-size:11px;opacity:.6;margin-top:3px;font-family:'IBM Plex Mono',monospace}}
.page-hdr .ref{{font-size:12px;background:rgba(255,255,255,.15);padding:5px 14px;border-radius:16px}}
.yield-banner{{display:flex;align-items:center;gap:20px;flex-wrap:wrap;
  background:#ecfdf5;border:2px solid #6ee7b7;margin:18px 36px;
  padding:14px 22px;border-radius:10px}}
.y-stat{{text-align:center}}
.y-val{{font-size:20px;font-weight:700;font-family:'IBM Plex Mono',monospace}}
.y-lbl{{font-size:10px;color:var(--muted);margin-top:2px;text-transform:uppercase;letter-spacing:.5px}}
.y-arrow{{font-size:22px;color:var(--accent);font-weight:700}}
.content{{padding:18px 36px 60px}}

/* RM Table */
.section-lbl{{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;
  color:var(--muted);border-left:4px solid var(--accent);padding-left:10px;
  margin:20px 0 10px}}
.rm-tbl{{width:100%;border-collapse:collapse;font-size:11px;background:white;
  border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.07);margin-bottom:28px}}
.rm-tbl th{{background:var(--navy);color:white;padding:8px 10px;text-align:left;
  font-size:10px;text-transform:uppercase;letter-spacing:.4px}}
.rm-tbl td{{padding:7px 10px;border-bottom:1px solid #e8f0f8}}
.rm-tbl .row-sm{{background:#dbeafe}}
.rm-tbl .row-product{{background:#d1fae5}}
.rm-tbl .row-solvent{{background:#f0fdf4}}
.rm-tbl .row-byproduct{{background:#fff7ed}}

/* Flow */
.bfd-flow{{display:flex;flex-direction:column;align-items:center;gap:0}}
.flow-node{{display:grid;grid-template-columns:200px 28px 1fr 28px 260px;
  width:100%;max-width:1140px;align-items:start;margin-bottom:4px}}
.inlet-col,.outlet-col{{display:flex;flex-direction:column;gap:6px;padding-top:10px}}
.arr-col{{display:flex;align-items:flex-start;justify-content:center;
  padding-top:18px;font-size:22px;color:var(--accent);font-weight:700}}

/* Stream cards */
.sc{{border-radius:7px;padding:8px 10px;font-size:11px;
  border:1.5px solid rgba(0,0,0,.12);box-shadow:0 1px 4px rgba(0,0,0,.06)}}
.sc .s-id{{font-family:'IBM Plex Mono',monospace;font-size:10px;
  font-weight:600;margin-bottom:2px}}
.sc .s-name{{font-weight:600;margin-bottom:2px}}
.sc .s-qty{{font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--muted)}}
.sc-inlet{{background:#eff6ff;border-color:#bfdbfe}}
.sc .s-note{{font-size:9px;opacity:.65;margin-top:2px;font-style:italic}}

/* Unit block */
.unit-block{{background:var(--block);border:2px solid var(--border);
  border-radius:10px;overflow:hidden;box-shadow:0 4px 14px rgba(30,90,160,.13)}}
.blk-hdr{{background:var(--blue);color:white;padding:10px 16px;
  font-size:13px;font-weight:700;letter-spacing:.3px}}
.blk-body{{padding:10px 14px}}
.step-list{{list-style:none;margin-bottom:8px}}
.step-list li{{font-size:11px;padding:3px 0;border-bottom:1px solid rgba(0,0,0,.07);
  display:flex;gap:7px}}
.step-list li:last-child{{border-bottom:none}}
.step-num{{min-width:18px;height:18px;background:var(--blue);color:white;
  border-radius:50%;display:flex;align-items:center;justify-content:center;
  font-size:9px;font-weight:700;flex-shrink:0;margin-top:1px}}
.eq-grid{{display:grid;grid-template-columns:1fr 1fr;gap:3px;
  font-size:10px;background:rgba(0,0,0,.04);padding:7px;border-radius:6px;margin-top:5px}}
.eq-lbl{{font-size:9px;text-transform:uppercase;color:var(--muted)}}
.eq-val{{font-family:'IBM Plex Mono',monospace;font-weight:600;font-size:10px}}
.warn-box{{background:#fffbeb;border:1.5px solid #f59e0b;border-radius:6px;
  padding:6px 10px;font-size:10px;color:#92400e;margin-top:6px}}

/* T/P badge */
.tp-badge{{background:white;border:1.5px solid var(--border);border-radius:7px;
  padding:7px 10px;font-size:10px;margin-top:6px}}
.tp-badge .tp-lbl{{font-size:9px;text-transform:uppercase;color:var(--muted);margin-bottom:3px}}

/* Process stream */
.ps-wrap{{display:flex;flex-direction:column;align-items:center;
  width:100%;max-width:1140px;margin:0}}
.ps-arrow{{font-size:26px;color:var(--accent);line-height:1}}
.ps-card{{background:var(--ps-bg);border:2px solid var(--ps-border);
  border-radius:8px;width:100%;padding:11px 16px;font-size:12px;
  box-shadow:0 2px 10px rgba(92,184,92,.15)}}
.ps-hdr{{display:flex;justify-content:space-between;align-items:center;
  font-family:'IBM Plex Mono',monospace;font-weight:700;font-size:12px;
  color:var(--navy);margin-bottom:7px}}
.ps-totals{{font-size:11px;color:var(--muted)}}
.comp-tbl{{width:100%;border-collapse:collapse;font-size:11px;margin-top:4px}}
.comp-tbl th{{background:rgba(30,90,160,.1);padding:4px 8px;
  text-align:left;font-size:10px;text-transform:uppercase}}
.comp-tbl td{{padding:3px 8px;border-bottom:1px solid rgba(0,0,0,.06)}}
.comp-tbl tr:last-child td{{border-bottom:none}}
.next-arrow{{width:100%;max-width:1140px;display:flex;justify-content:center;
  padding:10px 0;font-size:22px;color:var(--accent)}}

/* Legend */
.legend{{display:flex;gap:12px;flex-wrap:wrap;margin:0 0 20px}}
.leg-item{{display:flex;align-items:center;gap:5px;font-size:11px}}
.leg-dot{{width:14px;height:14px;border-radius:3px;border:1px solid rgba(0,0,0,.15)}}</style>
</head>
<body>
<div class="page-hdr">
  <div>
    <h1>BLOCK FLOW DIAGRAM — {product_name}</h1>
    <div class="meta">Mass Balance Report{(" | Ref: "+reference) if reference else ""}</div>
  </div>
  <div class="ref">Batch: {batch_size} kg SM</div>
</div>
{yield_html}
<div class="content">
  <div class="legend">
    <div class="leg-item"><div class="leg-dot" style="background:#eff6ff;border-color:#bfdbfe"></div>Inlet Stream (IS)</div>
    <div class="leg-item"><div class="leg-dot" style="background:#e2efda;border-color:#5cb85c"></div>Process Stream (PS)</div>
    <div class="leg-item"><div class="leg-dot" style="background:#dbeafe;border-color:#93c5fd"></div>Gas Waste</div>
    <div class="leg-item"><div class="leg-dot" style="background:#fef9c3;border-color:#fde047"></div>Organic Waste</div>
    <div class="leg-item"><div class="leg-dot" style="background:#fce7f3;border-color:#f9a8d4"></div>Aqueous Waste</div>
    <div class="leg-item"><div class="leg-dot" style="background:#ffedd5;border-color:#fdba74"></div>Solid Waste</div>
  </div>

  <div class="section-lbl">Raw Material Table</div>
  <table class="rm-tbl">
    <thead><tr>
      <th>Component</th><th>MW</th><th>Mol. Ratio</th><th>Moles (kmol)</th>
      <th>Purity</th><th>Qty (kg)</th><th>Density (kg/m³)</th><th>Volume (kL)</th>
    </tr></thead>
    <tbody>{rm_rows}</tbody>
  </table>

  <div class="section-lbl">Block Flow Diagram</div>
  <div class="bfd-flow">
    {blocks_html}
  </div>
</div>
</body></html>"""
    return html


def _render_block(op):
    op_name  = op.get("name", op.get("type","")).upper()
    op_type  = op.get("type","")
    eq       = op.get("equipment", {})
    steps    = op.get("steps", [])
    inlets   = op.get("inlet_streams", [])
    outlets  = op.get("outlet_streams", [])
    ps       = op.get("process_stream")
    t_init   = op.get("temp_initial","—")
    t_final  = op.get("temp_final","—")
    pressure = op.get("pressure","atm")

    # Inlet cards
    inlets_html = ""
    for s in inlets:
        inlets_html += f"""
        <div class="sc sc-inlet">
          <div class="s-id">{s['id']}</div>
          <div class="s-name">{s['component']}</div>
          <div class="s-qty">{s.get('qty_kg','—')} kg · {s.get('vol_kl','—')} kL</div>
        </div>"""

    # Outlet cards
    outlets_html = ""
    for s in outlets:
        cat        = s.get("category","organic_waste")
        bg, border_c, label = WASTE_COLORS.get(cat, ("#fef9c3","#854d0e","Waste"))
        note_html  = f'<div class="s-note">{s["note"]}</div>' if s.get("note") else ""
        outlets_html += f"""
        <div class="sc" style="background:{bg};border-color:{border_c}">
          <div class="s-id" style="color:{border_c}">{s['id']} · {label}</div>
          <div class="s-name">{s['component']}</div>
          <div class="s-qty">{s.get('qty_kg','—')} kg · {s.get('vol_kl','—')} kL</div>
          {note_html}
        </div>"""

    # T/P badge
    outlets_html += f"""
    <div class="tp-badge">
      <div class="tp-lbl">Temperature / Pressure</div>
      <div style="font-size:10px">T: {t_init} → {t_final} °C &nbsp;|&nbsp; P: {pressure}</div>
    </div>"""

    # Volume warning
    if op.get("volume_warning"):
        outlets_html += f'<div class="warn-box">{op["volume_warning"]}</div>'

    # Steps
    if steps:
        steps_html = '<ul class="step-list">'
        for i, s in enumerate(steps, 1):
            steps_html += f'<li><span class="step-num">{i}</span><span>{s}</span></li>'
        steps_html += "</ul>"
    else:
        steps_html = f'<div style="font-size:11px;color:var(--muted);padding:8px 0">[{op_name} steps]</div>'

    eq_html = f"""<div class="eq-grid">
      <div><div class="eq-lbl">Tag</div><div class="eq-val">{eq.get('tag','—')}</div></div>
      <div><div class="eq-lbl">MOC</div><div class="eq-val">{eq.get('moc','SS-316L')}</div></div>
      <div><div class="eq-lbl">Op. Vol (kL)</div><div class="eq-val">{op.get('operating_volume_kl','—')}</div></div>
      <div><div class="eq-lbl">Eq. Vol (kL)</div><div class="eq-val">{eq.get('volume_kl','—')}</div></div>
    </div>"""

    # Process stream
    ps_html = ""
    if ps:
        comp_rows = ""
        for c in ps.get("composition", []):
            comp_rows += f"""<tr>
              <td>{c.get('component','')}</td>
              <td style="font-family:'IBM Plex Mono',monospace">{c.get('qty_kg','')}</td>
              <td style="font-family:'IBM Plex Mono',monospace">{c.get('wt_pct','—')}%</td>
            </tr>"""
        comp_tbl = (f"""<table class="comp-tbl">
          <thead><tr><th>Component</th><th>Qty (kg)</th><th>wt%</th></tr></thead>
          <tbody>{comp_rows}</tbody>
        </table>""" if comp_rows else "")
        ps_html = f"""
      <div class="ps-wrap">
        <div class="ps-arrow">↓</div>
        <div class="ps-card">
          <div class="ps-hdr">
            <span>{ps['id']} — {ps['name']}</span>
            <span class="ps-totals">{ps.get('qty_kg','—')} kg · {ps.get('vol_kl','—')} kL</span>
          </div>
          {comp_tbl}
        </div>
      </div>
      <div class="next-arrow">↓</div>"""

    return f"""
    <div class="flow-node">
      <div class="inlet-col">{inlets_html}</div>
      <div class="arr-col">→</div>
      <div>
        <div class="unit-block">
          <div class="blk-hdr">UNIT OPERATIONS — {op_name}</div>
          <div class="blk-body">{steps_html}{eq_html}</div>
        </div>
      </div>
      <div class="arr-col">→</div>
      <div class="outlet-col">{outlets_html}</div>
    </div>
    {ps_html}"""
