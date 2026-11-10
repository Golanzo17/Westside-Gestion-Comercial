---
name: winforms-responsive-ui
description: Ensure all WinForms forms, modals, toolbars, KPI cards, SplitContainers, and DataGridViews in GestionComercial are 100% responsive across window resizing, maximizing, and smaller screen resolutions (1024px–1920px+). Activate whenever creating or modifying any UI form or control in VB.NET.
---

# WinForms Responsive UI Guidelines (`GestionComercial`)

Always apply these responsive layout rules whenever creating or modifying any `Form`, `Panel`, `DataGridView`, `SplitContainer`, `TableLayoutPanel`, `FlowLayoutPanel`, or modal dialog in the project.

## 1. DataGridViews (`AutoSizeColumnsMode.Fill` + `FillWeight` + `MinimumWidth`)
- **Never** use fixed pixel widths (`Width = 130`) across all columns of a `DataGridView`.
- Always set `dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill` and assign every visible column a proportional `FillWeight` and a sensible `MinimumWidth`:
  ```vb
  dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
  dgv.Columns.Add("Ticket", "Nº Ticket")
  dgv.Columns("Ticket").FillWeight = 14
  dgv.Columns("Ticket").MinimumWidth = 105
  ```

## 2. Toolbars and Filter Bars (`FlowLayoutPanel` or `Resize` Handler)
- Wrap horizontal filter/action bars inside a `Panel` (`Dock = DockStyle.Top`, `AutoSize = True`) containing a `FlowLayoutPanel` with:
  - `Dock = DockStyle.Top`
  - `AutoSize = True`
  - `AutoSizeMode = AutoSizeMode.GrowAndShrink`
  - `WrapContents = True`
- Alternatively, if using a multi-column filter panel (like `FrmControlCajas`), bind a proportional layout calculation to **both** `pnlFiltros.Resize` and `Me.Shown`.

## 3. KPI Cards (`TableLayoutPanel` + `UITheme.CreateKpiCard`)
- Always place KPI cards inside a `TableLayoutPanel` (`Dock = DockStyle.Fill`) with equal percentage columns (`SizeType.Percent`, e.g., `25.0F` each for 4 cards).
- Use `UITheme.CreateKpiCard(...)` with `Dock = DockStyle.Fill`. Ensure internal labels use `AutoSize = False`, `AutoEllipsis = True`, and resize their width dynamically on `pnl.Resize`.

## 4. `SplitContainer` Panels
- Never leave a `SplitContainer` with only a static `SplitterDistance` that clips one panel when the window shrinks.
- Set `Panel1MinSize` and `Panel2MinSize`, and attach a `Resize` handler that recalculates `SplitterDistance` proportionally (e.g., `CInt(sc.Height * 0.48)` or `CInt(sc.Width * 0.62)`) inside a `Try...Catch` block.

## 5. Labels, Inputs, and Buttons in Cards or Modals
- In detail modals and secondary dialogs (`MostrarModalDetalleVenta`, `BtnVerMovimientos`, `BtnVerTickets`, `FrmDetalleArqueoCaja`):
  - Prefer `FormBorderStyle.Sizable` with a sensible `MinimumSize` when displaying a `DataGridView`.
  - Anchor right-aligned footer buttons with `Anchor = AnchorStyles.Bottom Or AnchorStyles.Right` and position them relative to `ClientSize.Width`.
  - For key-value rows (such as cash reconciliation formulas), anchor the left description label with `AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right` and `AutoEllipsis = True`, and anchor the right value label with `AnchorStyles.Top Or AnchorStyles.Right`.

