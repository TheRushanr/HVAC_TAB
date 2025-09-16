"""
HVAC Testing & Air Balancing App (PyQt5) — with Auto-balance suggestions & Charts
Version: 1.0.4
Signature: RR23
Features:
- Procedure checklist (FCU/AHU/Extract/Terminals/Pressurization)
- Measurements table (add/update/remove)
- Optional fields: Vibration, Off-coil Temp, On-coil Temp
- Automatic ΔT calculation (On - Off)
- Auto-balance suggestions (suggested damper close %) to match index outlet %
- Embedded charts (matplotlib): measured vs target, cumulative flow
- Export CSV (always) and PDF (optional, reportlab)
Author: RR23
"""

import sys, csv
from datetime import datetime
from PyQt5 import QtWidgets, QtCore, QtGui

# Optional PDF library
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

# Matplotlib for charts
try:
    import matplotlib
    matplotlib.use("Qt5Agg")
    from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
    from matplotlib.figure import Figure
    MATPLOTLIB_AVAILABLE = True
except Exception:
    MATPLOTLIB_AVAILABLE = False

# Condensed method statement steps (from user's PDF)
BALANCING_STEPS = {
    "Fan Coil Unit (FCU)": [
        "Check automatic controls commissioned & operating",
        "Ensure pre-commissioning checks are complete",
        "Select specified fan speed",
        "Open all outlet dampers fully",
        "Take initial total flow (sum outlets)",
        "Compare vs design flow (compute %)",
        "Identify index outlet (lowest %)",
        "Keep index outlet damper fully open",
        "Throttle other outlets proportionally using flow hood",
        "Measure index each time until balanced",
        "Record final readings & save records"
    ],
    "Air Handling Unit (AHU)": [
        "Set fan RPM to provide design total air quantity",
        "Ensure fan current ≤ manufacturer limits",
        "Open all main & branch dampers fully",
        "Check total flow by traverse method (set to 105% of design)",
        "Identify index branch",
        "Balance branches proportionally using VCDs",
        "Record fan RPM, motor V/A, entering/leaving static pressures"
    ],
    "Ventilation / Extract Fan": [
        "Pre-commissioning checks complete",
        "Measure motor amperes & fan RPM",
        "Ensure speed/current within allowable range",
        "Open main & branch dampers fully",
        "Check total flow by traverse (105% of design)",
        "Identify index branch & balance branches",
        "Record index branch and final results"
    ],
    "Air Terminals (Diffusers/Grilles)": [
        "Measure flow at each outlet (flow hood preferred)",
        "Find index terminal (lowest percentage)",
        "Adjust other outlets proportionally to index",
        "Re-measure and record terminal flows",
        "Sum to check total vs branch measured flow"
    ],
    "Pressurization Fan (Staircase)": [
        "Pre-commissioning checks complete",
        "Ensure fire alarm interfacing verified",
        "Open all outlet dampers fully",
        "Take initial outlet readings (sum)",
        "Identify index outlet",
        "Keep index open, throttle others proportionally",
        "Balance outlets proportionally and record readings",
        "Measure differential pressure between floor & stair"
    ]
}

# ----------------------------
# Helper utilities
# ----------------------------
def safe_float(s, default=0.0):
    try:
        return float(s)
    except Exception:
        return default

def percent_of_design(flow, design_flow):
    if design_flow == 0:
        return 0.0
    return 100.0 * flow / design_flow

def compute_totals(measurements):
    return sum([safe_float(m.get("flow", 0.0)) for m in measurements])

# ----------------------------
# Export: CSV
# ----------------------------
def export_csv(filepath, meta, measurements, final_notes):
    with open(filepath, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["HVAC Testing & Air Balancing Report"])
        w.writerow(["Procedure Type", meta.get("type","")])
        w.writerow(["Date", meta.get("date","")])
        w.writerow(["Operator / Technician", meta.get("operator","")])
        w.writerow([])
        w.writerow(["Design Flow", f"{meta.get('design_flow','')} {meta.get('flow_unit','')}"])
        if meta.get("rpm"): w.writerow(["Design Fan RPM", meta.get("rpm")])
        if meta.get("voltage"): w.writerow(["Motor Voltage (V)", meta.get("voltage")])
        if meta.get("current"): w.writerow(["Motor Current (A)", meta.get("current")])
        if meta.get("entering_sp"): w.writerow(["Entering static pressure (Pa)", meta.get("entering_sp")])
        if meta.get("leaving_sp"): w.writerow(["Leaving static pressure (Pa)", meta.get("leaving_sp")])
        if meta.get("vibration"): w.writerow(["Vibration (mm/s)", meta.get("vibration")])
        if meta.get("off_coil"): w.writerow(["Off-coil Temp (°C)", meta.get("off_coil")])
        if meta.get("on_coil"): w.writerow(["On-coil Temp (°C)", meta.get("on_coil")])
        if "deltaT" in meta and meta["deltaT"] is not None:
            w.writerow(["ΔT (On - Off) (°C)", f"{meta['deltaT']:.2f}"])
        w.writerow([])
        w.writerow(["Outlet/Branch ID","Measured Flow (m3/s)","Measured Flow (orig unit)","% of Design","Static Pressure (Pa)","Note","Suggested damper close (%)"])
        for m in measurements:
            w.writerow([
                m.get("id",""),
                f"{m.get('flow',0.0):.6f}",
                m.get("orig_flow",""),
                f"{m.get('pct',0.0):.1f}",
                m.get("static",""),
                m.get("note",""),
                f"{m.get('suggested_close',0.0):.1f}"
            ])
        w.writerow([])
        w.writerow(["Total measured flow (m3/s)", f"{meta.get('total_m3s',0.0):.6f}"])
        w.writerow(["Total measured flow (unit)", meta.get("total_unit","")])
        w.writerow(["% of design (total)", f"{meta.get('total_pct',0.0):.1f}"])
        w.writerow([])
        w.writerow(["Final notes:"])
        w.writerow([final_notes])

# ----------------------------
# Export: PDF (reportlab, optional)
# ----------------------------
def export_pdf(filepath, meta, measurements, final_notes):
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("reportlab not installed")
    c = canvas.Canvas(filepath, pagesize=A4)
    w_page, h_page = A4
    margin = 18 * mm
    y = h_page - 20 * mm

    def writeline(text, size=9, dy=6*mm):
        nonlocal y
        c.setFont("Helvetica", size)
        c.drawString(margin, y, text)
        y -= dy
        if y < 40*mm:
            c.showPage()
            y = h_page - 20 * mm

    writeline("HVAC Testing & Air Balancing Report", size=12)
    writeline(f"Procedure Type: {meta.get('type','')}")
    writeline(f"Date: {meta.get('date','')}")
    writeline(f"Operator: {meta.get('operator','')}")
    writeline("")
    writeline(f"Design Flow: {meta.get('design_flow','')} {meta.get('flow_unit','')}")
    if meta.get("rpm"): writeline(f"Fan RPM: {meta.get('rpm')}")
    if meta.get("voltage"): writeline(f"Motor Voltage (V): {meta.get('voltage')}")
    if meta.get("current"): writeline(f"Motor Current (A): {meta.get('current')}")
    if meta.get("entering_sp"): writeline(f"Entering SP (Pa): {meta.get('entering_sp')}")
    if meta.get("leaving_sp"): writeline(f"Leaving SP (Pa): {meta.get('leaving_sp')}")
    if meta.get("vibration"): writeline(f"Vibration: {meta.get('vibration')}")
    if meta.get("off_coil"): writeline(f"Off-coil Temp (°C): {meta.get('off_coil')}")
    if meta.get("on_coil"): writeline(f"On-coil Temp (°C): {meta.get('on_coil')}")
    if "deltaT" in meta and meta["deltaT"] is not None:
        writeline(f"ΔT (On - Off) (°C): {meta['deltaT']:.2f}")
    writeline("")
    writeline("Measurements:", size=10)
    writeline("ID | Flow (m3/s) | Flow(orig unit) | % Design | Static (Pa) | Note | Suggested damper close (%)", size=8)
    for m in measurements:
        s = f"{m.get('id','')} | {m.get('flow',0.0):.6f} | {m.get('orig_flow','')} | {m.get('pct',0.0):.1f}% | {m.get('static','')} | {m.get('note','')} | {m.get('suggested_close',0.0):.1f}%"
        writeline(s, size=8)
    writeline("")
    writeline(f"Total measured flow (m3/s): {meta.get('total_m3s',0.0):.6f}")
    writeline(f"Total measured flow ({meta.get('flow_unit','')}): {meta.get('total_unit','')}")
    writeline(f"% of design (total): {meta.get('total_pct',0.0):.1f}%")
    writeline("")
    writeline("Final notes:", size=10)
    for line in final_notes.splitlines():
        writeline(line, size=9)
    c.save()

# ----------------------------
# Main Application UI
# ----------------------------
class BalancingApp(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("HVAC Testing & Air Balancing — Auto-balance + Charts")
        self.resize(1200, 820)
        self._measurements = []  # list of dict records
        self._last_summary = {}
        self._build_ui()

    def _build_ui(self):
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        vbox = QtWidgets.QVBoxLayout(central)

        # Top panels: selection/checklist, design info, measurements
        top_h = QtWidgets.QHBoxLayout()
        vbox.addLayout(top_h)

        # Left: procedure & checklist
        left_grp = QtWidgets.QGroupBox("Procedure & Checklist")
        left_layout = QtWidgets.QVBoxLayout(left_grp)
        top_h.addWidget(left_grp, stretch=2)

        self.type_combo = QtWidgets.QComboBox()
        self.type_combo.addItems(list(BALANCING_STEPS.keys()))
        self.type_combo.currentTextChanged.connect(self._load_steps)
        left_layout.addWidget(QtWidgets.QLabel("Select balancing type:"))
        left_layout.addWidget(self.type_combo)

        self.steps_list = QtWidgets.QListWidget()
        left_layout.addWidget(QtWidgets.QLabel("Checklist:"))
        left_layout.addWidget(self.steps_list)

        # Middle: design & test info (including vibration and coil temps)
        mid_grp = QtWidgets.QGroupBox("Design & Test Info")
        mid_form = QtWidgets.QFormLayout(mid_grp)
        top_h.addWidget(mid_grp, stretch=3)

        self.operator_edit = QtWidgets.QLineEdit()
        mid_form.addRow("Operator / Technician:", self.operator_edit)

        self.design_flow_edit = QtWidgets.QLineEdit()
        self.design_flow_edit.setPlaceholderText("e.g. 1000")
        mid_form.addRow("Design flow:", self.design_flow_edit)

        self.flow_unit_cb = QtWidgets.QComboBox()
        self.flow_unit_cb.addItems(["CFM", "m3/s", "m3/h", "L/s"])
        mid_form.addRow("Flow unit:", self.flow_unit_cb)

        self.rpm_edit = QtWidgets.QLineEdit(); mid_form.addRow("Fan RPM (optional):", self.rpm_edit)
        self.voltage_edit = QtWidgets.QLineEdit(); mid_form.addRow("Motor Voltage (V, optional):", self.voltage_edit)
        self.current_edit = QtWidgets.QLineEdit(); mid_form.addRow("Motor Current (A, optional):", self.current_edit)
        self.entering_sp_edit = QtWidgets.QLineEdit(); mid_form.addRow("Entering static pressure (Pa, optional):", self.entering_sp_edit)
        self.leaving_sp_edit = QtWidgets.QLineEdit(); mid_form.addRow("Leaving static pressure (Pa, optional):", self.leaving_sp_edit)

        # NEW optional fields
        self.vibration_edit = QtWidgets.QLineEdit(); mid_form.addRow("Vibration (mm/s) (optional):", self.vibration_edit)
        self.off_coil_edit = QtWidgets.QLineEdit(); mid_form.addRow("Off-coil Temp (°C) (optional):", self.off_coil_edit)
        self.on_coil_edit = QtWidgets.QLineEdit(); mid_form.addRow("On-coil Temp (°C) (optional):", self.on_coil_edit)

        self.notes_edit = QtWidgets.QPlainTextEdit()
        self.notes_edit.setPlaceholderText("Observations, corrective actions, remarks...")
        mid_form.addRow("Notes:", self.notes_edit)

        # Right: measurement inputs + table
        right_grp = QtWidgets.QGroupBox("Measurements")
        right_v = QtWidgets.QVBoxLayout(right_grp)
        top_h.addWidget(right_grp, stretch=5)

        meas_form = QtWidgets.QFormLayout()
        self.outlet_id = QtWidgets.QLineEdit()
        self.meas_flow = QtWidgets.QLineEdit()
        self.meas_flow.setPlaceholderText("Measured flow (in selected unit)")
        self.meas_static = QtWidgets.QLineEdit()
        self.meas_static.setPlaceholderText("Static pressure (Pa) optional")
        self.meas_note = QtWidgets.QLineEdit()
        meas_form.addRow("Outlet / Branch ID:", self.outlet_id)
        meas_form.addRow("Measured flow:", self.meas_flow)
        meas_form.addRow("Static pressure (Pa):", self.meas_static)
        meas_form.addRow("Note:", self.meas_note)
        right_v.addLayout(meas_form)

        btn_h = QtWidgets.QHBoxLayout()
        self.add_btn = QtWidgets.QPushButton("Add Measurement")
        self.add_btn.clicked.connect(self.add_measurement)
        self.update_btn = QtWidgets.QPushButton("Update Selected")
        self.update_btn.clicked.connect(self.update_selected)
        self.remove_btn = QtWidgets.QPushButton("Remove Selected")
        self.remove_btn.clicked.connect(self.remove_selected)
        btn_h.addWidget(self.add_btn); btn_h.addWidget(self.update_btn); btn_h.addWidget(self.remove_btn)
        right_v.addLayout(btn_h)

        # Table columns include suggested damper close %
        self.table = QtWidgets.QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels([
            "Outlet/Branch ID", "Flow (m3/s)", "Flow (orig unit)", "% of design",
            "Static (Pa)", "Note", "Suggested damper close (%)"
        ])
        self.table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        right_v.addWidget(self.table)
        self.table.itemSelectionChanged.connect(self.populate_from_selected)

        # Charts area (matplotlib) if available
        chart_grp = QtWidgets.QGroupBox("Charts")
        chart_v = QtWidgets.QVBoxLayout(chart_grp)
        vbox.addWidget(chart_grp, stretch=3)

        if MATPLOTLIB_AVAILABLE:
            self.fig = Figure(figsize=(8,3))
            self.canvas = FigureCanvas(self.fig)
            chart_v.addWidget(self.canvas)
            self.plot_btn = QtWidgets.QPushButton("Update Charts")
            self.plot_btn.clicked.connect(self.update_charts)
            chart_v.addWidget(self.plot_btn)
        else:
            label = QtWidgets.QLabel("matplotlib not installed — charts disabled. Install with: pip install matplotlib")
            chart_v.addWidget(label)

        # Bottom: summary and actions
        bottom_h = QtWidgets.QHBoxLayout()
        vbox.addLayout(bottom_h)

        summary_grp = QtWidgets.QGroupBox("Summary")
        summary_form = QtWidgets.QFormLayout(summary_grp)
        bottom_h.addWidget(summary_grp, stretch=3)

        self.total_flow_lbl = QtWidgets.QLabel("0.000000 m3/s")
        self.total_unit_lbl = QtWidgets.QLabel("")
        self.total_pct_lbl = QtWidgets.QLabel("0.0 %")
        self.index_outlet_lbl = QtWidgets.QLabel("-")
        self.deltaT_lbl = QtWidgets.QLabel("N/A")
        summary_form.addRow("Total measured flow (m3/s):", self.total_flow_lbl)
        summary_form.addRow("Total measured flow (unit):", self.total_unit_lbl)
        summary_form.addRow("% of design:", self.total_pct_lbl)
        summary_form.addRow("Index outlet (lowest %):", self.index_outlet_lbl)
        summary_form.addRow("ΔT (On - Off) (°C):", self.deltaT_lbl)

        actions_v = QtWidgets.QVBoxLayout()
        bottom_h.addLayout(actions_v, stretch=2)

        self.compute_btn = QtWidgets.QPushButton("Compute Totals / Find Index & ΔT / Suggest Dampers")
        self.compute_btn.clicked.connect(self.compute_and_update)
        actions_v.addWidget(self.compute_btn)

        self.export_csv_btn = QtWidgets.QPushButton("Export CSV Report")
        self.export_csv_btn.clicked.connect(self.export_csv)
        actions_v.addWidget(self.export_csv_btn)

        self.export_pdf_btn = QtWidgets.QPushButton("Export PDF Report (optional)")
        self.export_pdf_btn.clicked.connect(self.export_pdf)
        if not REPORTLAB_AVAILABLE:
            self.export_pdf_btn.setEnabled(False)
            self.export_pdf_btn.setToolTip("Install 'reportlab' to enable PDF export: pip install reportlab")
        actions_v.addWidget(self.export_pdf_btn)

        # initial checklist load
        self._load_steps(self.type_combo.currentText())

    # ----------------------------
    # Checklist loader
    # ----------------------------
    def _load_steps(self, key):
        self.steps_list.clear()
        for s in BALANCING_STEPS.get(key, []):
            item = QtWidgets.QListWidgetItem(s)
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
            item.setCheckState(QtCore.Qt.Unchecked)
            self.steps_list.addItem(item)

    # ----------------------------
    # Measurement operations
    # ----------------------------
    def _convert_input_to_m3s(self, val, unit):
        u = unit.lower()
        if u == "m3/s": return val
        if u == "m3/h": return val / 3600.0
        if u == "l/s": return val / 1000.0
        if u == "cfm": return val * 0.00047194745
        return val

    def _convert_m3s_to_unit(self, m3s, unit):
        u = unit.lower()
        if u == "m3/s": return m3s
        if u == "m3/h": return m3s * 3600.0
        if u == "l/s": return m3s * 1000.0
        if u == "cfm": return m3s / 0.00047194745
        return m3s

    def add_measurement(self):
        id_ = self.outlet_id.text().strip()
        raw_flow_text = self.meas_flow.text().strip()
        static_text = self.meas_static.text().strip()
        note = self.meas_note.text().strip()
        if id_ == "" or raw_flow_text == "":
            QtWidgets.QMessageBox.warning(self, "Input required", "Please enter Outlet/Branch ID and Measured Flow.")
            return
        try:
            raw_val = float(raw_flow_text)
        except Exception:
            QtWidgets.QMessageBox.warning(self, "Bad input", "Measured flow must be numeric.")
            return
        flow_unit = self.flow_unit_cb.currentText()
        q_m3s = self._convert_input_to_m3s(raw_val, flow_unit)
        rec = {"id": id_, "flow": q_m3s, "orig_flow": raw_flow_text, "pct": 0.0, "static": static_text, "note": note, "suggested_close": 0.0}
        self._measurements.append(rec)
        self._append_row(rec)
        self.outlet_id.clear(); self.meas_flow.clear(); self.meas_static.clear(); self.meas_note.clear()

    def _append_row(self, rec):
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QtWidgets.QTableWidgetItem(rec["id"]))
        self.table.setItem(r, 1, QtWidgets.QTableWidgetItem(f"{rec['flow']:.6f}"))
        self.table.setItem(r, 2, QtWidgets.QTableWidgetItem(str(rec["orig_flow"])))
        self.table.setItem(r, 3, QtWidgets.QTableWidgetItem(f"{rec['pct']:.1f}"))
        self.table.setItem(r, 4, QtWidgets.QTableWidgetItem(str(rec["static"])))
        self.table.setItem(r, 5, QtWidgets.QTableWidgetItem(rec["note"]))
        self.table.setItem(r, 6, QtWidgets.QTableWidgetItem(f"{rec['suggested_close']:.1f}"))

    def populate_from_selected(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            return
        r = rows[0].row()
        rec = self._measurements[r]
        self.outlet_id.setText(rec.get("id",""))
        self.meas_flow.setText(str(rec.get("orig_flow","")))
        self.meas_static.setText(str(rec.get("static","")))
        self.meas_note.setText(rec.get("note",""))

    def update_selected(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            QtWidgets.QMessageBox.information(self, "Select row", "Select a table row to update.")
            return
        r = rows[0].row()
        id_ = self.outlet_id.text().strip()
        raw_flow_text = self.meas_flow.text().strip()
        static_text = self.meas_static.text().strip()
        note = self.meas_note.text().strip()
        if id_ == "" or raw_flow_text == "":
            QtWidgets.QMessageBox.warning(self, "Input required", "Please enter Outlet/Branch ID and Measured Flow.")
            return
        try:
            raw_val = float(raw_flow_text)
        except Exception:
            QtWidgets.QMessageBox.warning(self, "Bad input", "Measured flow must be numeric.")
            return
        q_m3s = self._convert_input_to_m3s(raw_val, self.flow_unit_cb.currentText())
        rec = {"id": id_, "flow": q_m3s, "orig_flow": raw_flow_text, "pct": 0.0, "static": static_text, "note": note, "suggested_close": 0.0}
        self._measurements[r] = rec
        # refresh row
        for c, val in enumerate([rec["id"], f"{rec['flow']:.6f}", rec["orig_flow"], f"{rec['pct']:.1f}", rec["static"], rec["note"], f"{rec['suggested_close']:.1f}"]):
            self.table.setItem(r, c, QtWidgets.QTableWidgetItem(str(val)))
        self.outlet_id.clear(); self.meas_flow.clear(); self.meas_static.clear(); self.meas_note.clear()
        self.compute_and_update()

    def remove_selected(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            QtWidgets.QMessageBox.information(self, "Select row", "Select a table row to remove.")
            return
        r = rows[0].row()
        self.table.removeRow(r)
        del self._measurements[r]
        self.compute_and_update()

    # ----------------------------
    # Compute totals, index outlet, ΔT, suggested damper closures
    # ----------------------------
    def compute_and_update(self):
        if not self._measurements:
            # clear summary
            self.total_flow_lbl.setText("0.000000 m3/s")
            self.total_unit_lbl.setText("")
            self.total_pct_lbl.setText("0.0 %")
            self.index_outlet_lbl.setText("-")
            self.deltaT_lbl.setText("N/A")
            return

        # design
        design_raw = safe_float(self.design_flow_edit.text(), 0.0)
        design_m3s = self._convert_input_to_m3s(design_raw, self.flow_unit_cb.currentText())

        # total measured flow
        total_m3s = compute_totals(self._measurements)

        # compute per-record percent of design and find index (lowest %)
        lowest_pct = None
        index_id = "-"
        for rec in self._measurements:
            pct = percent_of_design(rec["flow"], design_m3s) if design_m3s > 0 else 0.0
            rec["pct"] = pct
            if lowest_pct is None or pct < lowest_pct:
                lowest_pct = pct
                index_id = rec.get("id", "-")

        # Determine auto-balance target: match index_pct (keep index open)
        target_pct = lowest_pct if lowest_pct is not None else 0.0
        target_flow_per_outlet = design_m3s * (target_pct / 100.0)

        # Calculate suggested damper close % for each outlet
        for rec in self._measurements:
            current = rec.get("flow", 0.0)
            if current <= 0:
                rec["suggested_close"] = 0.0
            else:
                if current <= target_flow_per_outlet:
                    rec["suggested_close"] = 0.0
                else:
                    # simple linear suggestion: close such that flow reduces to target_flow
                    suggested = max(0.0, 100.0 * (1.0 - (target_flow_per_outlet / current)))
                    rec["suggested_close"] = suggested

        # update table rows
        for r, rec in enumerate(self._measurements):
            self.table.setItem(r, 1, QtWidgets.QTableWidgetItem(f"{rec['flow']:.6f}"))
            self.table.setItem(r, 3, QtWidgets.QTableWidgetItem(f"{rec['pct']:.1f}"))
            self.table.setItem(r, 6, QtWidgets.QTableWidgetItem(f"{rec['suggested_close']:.1f}"))

        # update summary
        self.total_flow_lbl.setText(f"{total_m3s:.6f} m3/s")
        total_in_unit = self._convert_m3s_to_unit(total_m3s, self.flow_unit_cb.currentText())
        self.total_unit_lbl.setText(f"{total_in_unit:.3f} {self.flow_unit_cb.currentText()}")
        total_pct = percent_of_design(total_m3s, design_m3s) if design_m3s > 0 else 0.0
        self.total_pct_lbl.setText(f"{total_pct:.1f} %")
        self.index_outlet_lbl.setText(index_id if index_id else "-")

        # ΔT calculation (On - Off)
        on_coil = None if self.on_coil_edit.text().strip() == "" else safe_float(self.on_coil_edit.text().strip(), None)
        off_coil = None if self.off_coil_edit.text().strip() == "" else safe_float(self.off_coil_edit.text().strip(), None)
        if on_coil is not None and off_coil is not None:
            deltaT = on_coil - off_coil
            self.deltaT_lbl.setText(f"{deltaT:.2f} °C")
        else:
            deltaT = None
            self.deltaT_lbl.setText("N/A")

        # store summary
        self._last_summary = {
            "design_m3s": design_m3s,
            "total_m3s": total_m3s,
            "total_in_unit": total_in_unit,
            "total_pct": total_pct,
            "index_id": index_id,
            "deltaT": deltaT,
            "target_pct": target_pct
        }

        # update charts
        if MATPLOTLIB_AVAILABLE:
            self.update_charts()

    # ----------------------------
    # Charts (matplotlib)
    # ----------------------------
    def update_charts(self):
        if not MATPLOTLIB_AVAILABLE:
            return
        self.fig.clear()
        ax1 = self.fig.add_subplot(1,2,1)
        ax2 = self.fig.add_subplot(1,2,2)

        # prepare data
        ids = [rec["id"] for rec in self._measurements]
        flows_m3s = [rec["flow"] for rec in self._measurements]
        unit = self.flow_unit_cb.currentText()
        flows_unit = [self._convert_m3s_to_unit(f, unit) for f in flows_m3s]

        # target per outlet based on index %
        design_raw = safe_float(self.design_flow_edit.text(), 0.0)
        design_m3s = self._convert_input_to_m3s(design_raw, unit)
        target_pct = self._last_summary.get("target_pct", 0.0)
        target_m3s = design_m3s * (target_pct / 100.0)
        targets_unit = [self._convert_m3s_to_unit(target_m3s, unit)] * len(ids)

        # Bar chart measured vs target (in chosen unit)
        ax1.bar([f"{i}\n(meas)" for i in ids], flows_unit, label="Measured")
        ax1.plot([f"{i}\n(meas)" for i in ids], targets_unit, 'r--', label=f"Target (index % = {target_pct:.1f}%)")
        ax1.set_title(f"Measured flow vs Target (unit: {unit})")
        ax1.set_ylabel(f"Flow ({unit})")
        ax1.legend()
        ax1.tick_params(axis='x', rotation=20)

        # Cumulative flow chart (measured)
        cum = []
        s = 0.0
        unit_vals = []
        for f in flows_m3s:
            s += f
            cum.append(s)
            unit_vals.append(self._convert_m3s_to_unit(s, unit))
        ax2.plot(ids, unit_vals, marker='o')
        ax2.set_title("Cumulative Measured Flow along branch")
        ax2.set_ylabel(f"Cumulative Flow ({unit})")
        ax2.grid(True)

        self.fig.tight_layout()
        self.canvas.draw()

    # ----------------------------
    # Collect meta and export helpers
    # ----------------------------
    def collect_meta(self):
        meta = {
            "type": self.type_combo.currentText(),
            "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "operator": self.operator_edit.text().strip(),
            "design_flow": self.design_flow_edit.text().strip(),
            "flow_unit": self.flow_unit_cb.currentText(),
            "rpm": self.rpm_edit.text().strip(),
            "voltage": self.voltage_edit.text().strip(),
            "current": self.current_edit.text().strip(),
            "entering_sp": self.entering_sp_edit.text().strip(),
            "leaving_sp": self.leaving_sp_edit.text().strip(),
            "vibration": self.vibration_edit.text().strip(),
            "off_coil": self.off_coil_edit.text().strip(),
            "on_coil": self.on_coil_edit.text().strip()
        }
        # ensure compute is up to date
        self.compute_and_update()
        meta["total_m3s"] = self._last_summary.get("total_m3s", 0.0)
        meta["total_unit"] = f"{self._last_summary.get('total_in_unit','')}"
        meta["total_pct"] = self._last_summary.get("total_pct", 0.0)
        meta["deltaT"] = self._last_summary.get("deltaT", None)
        return meta

    def export_csv(self):
        if not self._measurements:
            QtWidgets.QMessageBox.information(self, "No data", "Add measurements before exporting.")
            return
        fname, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save CSV", f"air_balance_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", "CSV files (*.csv);;All files (*)")
        if not fname:
            return
        meta = self.collect_meta()
        try:
            export_csv(fname, meta, self._measurements, self.notes_edit.toPlainText())
            QtWidgets.QMessageBox.information(self, "Saved", f"CSV exported to:\n{fname}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Export error", f"Failed to export CSV:\n{e}")

    def export_pdf(self):
        if not REPORTLAB_AVAILABLE:
            QtWidgets.QMessageBox.critical(self, "Missing dependency", "Install 'reportlab' to enable PDF export: pip install reportlab")
            return
        if not self._measurements:
            QtWidgets.QMessageBox.information(self, "No data", "Add measurements before exporting.")
            return
        fname, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save PDF", f"air_balance_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", "PDF files (*.pdf);;All files (*)")
        if not fname:
            return
        meta = self.collect_meta()
        try:
            export_pdf(fname, meta, self._measurements, self.notes_edit.toPlainText())
            QtWidgets.QMessageBox.information(self, "Saved", f"PDF exported to:\n{fname}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Export error", f"Failed to export PDF:\n{e}")

# ----------------------------
# Run app
# ----------------------------
def main():
    app = QtWidgets.QApplication(sys.argv)
    w = BalancingApp()
    w.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
