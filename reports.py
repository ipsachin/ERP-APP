# ============================================================
# reports.py
# PDF / report engine for Liquimech ERP Desktop App
# ============================================================

from __future__ import annotations

from pathlib import Path
from importlib import import_module
from typing import List, Optional

from tkinter import filedialog, simpledialog

from app_config import AppConfig

colors = None
A4 = None
getSampleStyleSheet = None
ParagraphStyle = None
mm = None
SimpleDocTemplate = None
Table = None
TableStyle = None
Paragraph = None
Spacer = None


def norm_text(value) -> str:
    return str(value or "").strip()


def safe_filename(text: str, replacement: str = "_") -> str:
    text = str(text or "").strip()
    invalid = '<>:"/\\|?*'
    for ch in invalid:
        text = text.replace(ch, replacement)
    text = "".join(c for c in text if ord(c) >= 32)
    while "__" in text:
        text = text.replace("__", "_")
    text = text.strip(" ._")
    return text or "output"


def get_reportlab():
    try:
        colors = import_module("reportlab.lib.colors")
        A4 = import_module("reportlab.lib.pagesizes").A4
        styles = import_module("reportlab.lib.styles")
        mm = import_module("reportlab.lib.units").mm
        platypus = import_module("reportlab.platypus")
        return {
            "colors": colors,
            "A4": A4,
            "getSampleStyleSheet": styles.getSampleStyleSheet,
            "ParagraphStyle": styles.ParagraphStyle,
            "mm": mm,
            "SimpleDocTemplate": platypus.SimpleDocTemplate,
            "Table": platypus.Table,
            "TableStyle": platypus.TableStyle,
            "Paragraph": platypus.Paragraph,
            "Spacer": platypus.Spacer,
        }
    except ImportError as exc:
        raise RuntimeError(
            "PDF reporting is not available in this build because ReportLab is missing."
        ) from exc


def ensure_reportlab_loaded() -> None:
    global colors, A4, getSampleStyleSheet, ParagraphStyle, mm
    global SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    if Paragraph is not None:
        return
    modules = get_reportlab()
    colors = modules["colors"]
    A4 = modules["A4"]
    getSampleStyleSheet = modules["getSampleStyleSheet"]
    ParagraphStyle = modules["ParagraphStyle"]
    mm = modules["mm"]
    SimpleDocTemplate = modules["SimpleDocTemplate"]
    Table = modules["Table"]
    TableStyle = modules["TableStyle"]
    Paragraph = modules["Paragraph"]
    Spacer = modules["Spacer"]


class ReportService:
    def __init__(self, service_hub, mailer=None):
        self.services = service_hub
        self.mailer = mailer

    # ========================================================
    # Public dialog helpers
    # ========================================================

    def generate_module_report_dialog(self, module_code: str) -> Optional[str]:
        AppConfig.ensure_directories()
        default_name = f"{safe_filename(module_code)}_Module_Report.pdf"
        pdf_path = filedialog.asksaveasfilename(
            title="Save Module PDF Report",
            defaultextension=".pdf",
            initialdir=str(AppConfig.EXPORTS_DIR),
            initialfile=default_name,
            filetypes=[("PDF files", "*.pdf")],
        )
        if not pdf_path:
            return None

        self.build_module_pdf(module_code, pdf_path)
        return pdf_path

    def generate_product_quote_dialog(self, product_code: str) -> Optional[str]:
        AppConfig.ensure_directories()
        default_name = f"{safe_filename(product_code)}_Product_Quote.pdf"
        pdf_path = filedialog.asksaveasfilename(
            title="Save Product Quote PDF",
            defaultextension=".pdf",
            initialdir=str(AppConfig.EXPORTS_DIR),
            initialfile=default_name,
            filetypes=[("PDF files", "*.pdf")],
        )
        if not pdf_path:
            return None

        self.build_product_quote_pdf(product_code, pdf_path)
        return pdf_path

    def generate_project_report_dialog(self, project_code: str) -> Optional[str]:
        AppConfig.ensure_directories()
        default_name = f"{safe_filename(project_code)}_Project_Report.pdf"
        pdf_path = filedialog.asksaveasfilename(
            title="Save Project PDF Report",
            defaultextension=".pdf",
            initialdir=str(AppConfig.EXPORTS_DIR),
            initialfile=default_name,
            filetypes=[("PDF files", "*.pdf")],
        )
        if not pdf_path:
            return None

        self.build_project_report_pdf(project_code, pdf_path)
        return pdf_path

    def email_module_report_dialog(self, module_code: str) -> Optional[str]:
        if not self.mailer:
            raise RuntimeError("Mailer service is not configured.")

        to_email = simpledialog.askstring("Email Module Report", "Enter recipient email address:")
        if not to_email:
            return None

        AppConfig.ensure_directories()
        pdf_path = str(AppConfig.TEMP_DIR / f"{safe_filename(module_code)}_Module_Report.pdf")
        self.build_module_pdf(module_code, pdf_path)

        subject = f"Liquimech Module Report - {module_code}"
        body = (
            f"Hello,\n\n"
            f"Please find attached the module report for:\n{module_code}\n\n"
            f"Regards,\nLiquimech ERP Desktop Pro"
        )
        self.mailer.send_email_with_attachment(to_email, subject, body, pdf_path)
        return pdf_path

    def email_product_quote_dialog(self, product_code: str) -> Optional[str]:
        if not self.mailer:
            raise RuntimeError("Mailer service is not configured.")

        to_email = simpledialog.askstring("Email Product Quote", "Enter recipient email address:")
        if not to_email:
            return None

        AppConfig.ensure_directories()
        pdf_path = str(AppConfig.TEMP_DIR / f"{safe_filename(product_code)}_Product_Quote.pdf")
        self.build_product_quote_pdf(product_code, pdf_path)

        subject = f"Liquimech Product Quote - {product_code}"
        body = (
            f"Hello,\n\n"
            f"Please find attached the product quote for:\n{product_code}\n\n"
            f"Regards,\nLiquimech ERP Desktop Pro"
        )
        self.mailer.send_email_with_attachment(to_email, subject, body, pdf_path)
        return pdf_path

    def email_project_report_dialog(self, project_code: str) -> Optional[str]:
        if not self.mailer:
            raise RuntimeError("Mailer service is not configured.")

        to_email = simpledialog.askstring("Email Project Report", "Enter recipient email address:")
        if not to_email:
            return None

        AppConfig.ensure_directories()
        pdf_path = str(AppConfig.TEMP_DIR / f"{safe_filename(project_code)}_Project_Report.pdf")
        self.build_project_report_pdf(project_code, pdf_path)

        subject = f"Liquimech Project Report - {project_code}"
        body = (
            f"Hello,\n\n"
            f"Please find attached the live project report for:\n{project_code}\n\n"
            f"Regards,\nLiquimech ERP Desktop Pro"
        )
        self.mailer.send_email_with_attachment(to_email, subject, body, pdf_path)
        return pdf_path

    # ========================================================
    # Core PDF builders
    # ========================================================

    def build_module_pdf(self, module_code: str, pdf_path: str) -> None:
        ensure_reportlab_loaded()
        bundle = self.services.modules.get_module_bundle(module_code)
        if not bundle.module:
            raise ValueError("Module not found.")

        module = bundle.module
        tasks = bundle.tasks or []
        components = bundle.components or []
        documents = bundle.documents or []

        total_task_hours = sum(float(t.estimated_hours or 0.0) for t in tasks)

        doc = self._doc(pdf_path)
        title_style, section_style, normal_style, small_bold = self._styles()
        elements = []

        elements.append(Paragraph("Liquimech Assembly Report", title_style))
        elements.append(Spacer(1, 4))

        summary_data = [
            ["Assembly Code", module.module_code],
            ["Quote Ref", module.quote_ref],
            ["Assembly Name", module.module_name],
            ["Description", module.description],
            ["Status", module.status],
            ["Header Est. Hours", f"{float(module.estimated_hours or 0.0):.2f}"],
            ["Task Rollup Hours", f"{total_task_hours:.2f}"],
            ["Assembly Stock On Hand", f"{float(module.stock_on_hand or 0.0):.2f}"],
            ["Documents", str(len(documents))],
        ]
        elements.append(self._summary_table(summary_data))
        elements.append(Spacer(1, 10))

        elements.append(Paragraph("Instruction / Build Notes", section_style))
        elements.append(Paragraph(module.instruction_text or "-", normal_style))
        elements.append(Spacer(1, 10))

        elements.append(Paragraph("Tasks / Subtasks", section_style))
        if tasks:
            task_rows = [[
                Paragraph("<b>Task</b>", small_bold),
                Paragraph("<b>Department</b>", small_bold),
                Paragraph("<b>Hours</b>", small_bold),
                Paragraph("<b>Status</b>", small_bold),
                Paragraph("<b>Dependency</b>", small_bold),
                Paragraph("<b>Notes</b>", small_bold),
            ]]

            task_lookup = {t.task_id: t for t in tasks}
            for t in tasks:
                dep_name = ""
                if norm_text(t.dependency_task_id) and t.dependency_task_id in task_lookup:
                    dep_name = task_lookup[t.dependency_task_id].task_name
                elif norm_text(t.dependency_task_id):
                    dep_name = t.dependency_task_id

                indent = "&nbsp;&nbsp;&nbsp;&nbsp;" if norm_text(t.parent_task_id) else ""
                task_rows.append([
                    Paragraph(f"{indent}{t.task_name}", normal_style),
                    Paragraph(t.department or "-", normal_style),
                    Paragraph(f"{float(t.estimated_hours or 0.0):.2f}", normal_style),
                    Paragraph(t.status or "-", normal_style),
                    Paragraph(dep_name or "-", normal_style),
                    Paragraph(t.notes or "-", normal_style),
                ])

            table = Table(
                task_rows,
                colWidths=[55 * mm, 28 * mm, 16 * mm, 24 * mm, 32 * mm, 35 * mm],
                repeatRows=1,
            )
            table.setStyle(self._table_style())
            elements.append(table)
        else:
            elements.append(Paragraph("No tasks found for this module.", normal_style))

        elements.append(Spacer(1, 10))
        elements.append(Paragraph("Components / BOM", section_style))
        if components:
            rows = [[
                Paragraph("<b>Component</b>", small_bold),
                Paragraph("<b>Qty</b>", small_bold),
                Paragraph("<b>SOH</b>", small_bold),
                Paragraph("<b>Supplier</b>", small_bold),
                Paragraph("<b>Lead</b>", small_bold),
                Paragraph("<b>Part Number</b>", small_bold),
                Paragraph("<b>Notes</b>", small_bold),
            ]]
            for c in components:
                rows.append([
                    Paragraph(c.component_name or "-", normal_style),
                    Paragraph(str(c.qty), normal_style),
                    Paragraph(str(c.soh_qty), normal_style),
                    Paragraph(c.preferred_supplier or "-", normal_style),
                    Paragraph(str(c.lead_time_days), normal_style),
                    Paragraph(c.part_number or "-", normal_style),
                    Paragraph(c.notes or "-", normal_style),
                ])

            table = Table(
                rows,
                colWidths=[40 * mm, 12 * mm, 12 * mm, 30 * mm, 12 * mm, 30 * mm, 42 * mm],
                repeatRows=1,
            )
            table.setStyle(self._table_style())
            elements.append(table)
        else:
            elements.append(Paragraph("No components found for this module.", normal_style))

        elements.append(Spacer(1, 10))
        elements.append(Paragraph("Documents", section_style))
        if documents:
            rows = [[
                Paragraph("<b>Section</b>", small_bold),
                Paragraph("<b>Document</b>", small_bold),
                Paragraph("<b>Type</b>", small_bold),
                Paragraph("<b>Path</b>", small_bold),
            ]]
            for d in documents:
                rows.append([
                    Paragraph(d.section_name or "-", normal_style),
                    Paragraph(d.doc_name or "-", normal_style),
                    Paragraph(d.doc_type or "-", normal_style),
                    Paragraph(d.file_path or "-", normal_style),
                ])

            table = Table(
                rows,
                colWidths=[25 * mm, 40 * mm, 28 * mm, 85 * mm],
                repeatRows=1,
            )
            table.setStyle(self._table_style())
            elements.append(table)
        else:
            elements.append(Paragraph("No documents found for this module.", normal_style))

        doc.build(elements)

    def build_product_quote_pdf(self, product_code: str, pdf_path: str) -> None:
        ensure_reportlab_loaded()
        bundle = self.services.products.get_product_bundle(product_code)
        if not bundle.product:
            raise ValueError("Product not found.")

        product = bundle.product
        links = bundle.module_links or []
        docs = bundle.product_documents or []
        workorders = bundle.workorders or []
        tasks_by_module = bundle.tasks_by_module or {}
        modules = {m.module_code: m for m in (bundle.modules or [])}
        product_parts = []
        try:
            product_parts = self.services.products.get_product_components(product_code)
        except Exception:
            product_parts = []

        doc = self._doc(pdf_path)
        title_style, section_style, normal_style, small_bold = self._styles()
        elements = []

        elements.append(Paragraph("Liquimech Product Quote Pack", title_style))
        elements.append(Spacer(1, 4))

        summary_data = [
            ["Product Code", product.product_code],
            ["Quote Ref", product.quote_ref],
            ["Product Name", product.product_name],
            ["Description", product.description],
            ["Revision", product.revision],
            ["Status", product.status],
            ["Assigned Assemblies", str(len(links))],
            ["Direct Product Parts", str(len(product_parts))],
            ["Aggregated Labour Hours", f"{float(bundle.total_hours or 0.0):.2f}"],
            ["Direct Product Parts Cost", f"{sum(float(c.qty or 0.0) * float(c.unit_price or 0.0) for c in product_parts):.2f}"],
        ]
        elements.append(self._summary_table(summary_data))
        elements.append(Spacer(1, 10))

        elements.append(Paragraph("Assembly Quote Structure", section_style))
        if links:
            rows = [[
                Paragraph("<b>Order</b>", small_bold),
                Paragraph("<b>Assembly Code</b>", small_bold),
                Paragraph("<b>Assembly Name</b>", small_bold),
                Paragraph("<b>Qty</b>", small_bold),
                Paragraph("<b>Dependency</b>", small_bold),
                Paragraph("<b>Description</b>", small_bold),
                Paragraph("<b>Total Hours</b>", small_bold),
            ]]

            for link in links:
                module = modules.get(link.module_code)
                module_name = module.module_name if module else ""
                module_desc = module.description if module else ""
                module_hours = sum(float(t.estimated_hours or 0.0) for t in tasks_by_module.get(link.module_code, []))
                rows.append([
                    Paragraph(str(link.module_order), normal_style),
                    Paragraph(link.module_code, normal_style),
                    Paragraph(module_name or "-", normal_style),
                    Paragraph(str(link.module_qty), normal_style),
                    Paragraph(link.dependency_module_code or "-", normal_style),
                    Paragraph(module_desc or "-", normal_style),
                    Paragraph(f"{module_hours * link.module_qty:.2f}", normal_style),
                ])

            table = Table(
                rows,
                colWidths=[12 * mm, 35 * mm, 35 * mm, 12 * mm, 30 * mm, 48 * mm, 18 * mm],
                repeatRows=1,
            )
            table.setStyle(self._table_style())
            elements.append(table)
        else:
            elements.append(Paragraph("No modules assigned to this product.", normal_style))

        elements.append(Spacer(1, 10))
        elements.append(Paragraph("Detailed Assembly Task Breakdown", section_style))
        if links:
            for link in links:
                module = modules.get(link.module_code)
                module_name = module.module_name if module else ""
                module_desc = module.description if module else ""
                module_tasks = tasks_by_module.get(link.module_code, [])
                module_hours = sum(float(t.estimated_hours or 0.0) for t in module_tasks)

                elements.append(
                    Paragraph(
                        f"{link.module_order:02d}. {link.module_code} - {module_name} (Qty {link.module_qty})",
                        small_bold
                    )
                )

                info_rows = [
                    ["Description", module_desc or "-"],
                    ["Dependency", link.dependency_module_code or "-"],
                    ["Hours / Module", f"{module_hours:.2f}"],
                    ["Hours / Total Qty", f"{module_hours * link.module_qty:.2f}"],
                    ["Notes", link.notes or "-"],
                ]
                elements.append(self._summary_table(info_rows))
                elements.append(Spacer(1, 4))

                if module_tasks:
                    task_rows = [[
                        Paragraph("<b>Task</b>", small_bold),
                        Paragraph("<b>Department</b>", small_bold),
                        Paragraph("<b>Hours</b>", small_bold),
                        Paragraph("<b>Status</b>", small_bold),
                        Paragraph("<b>Notes</b>", small_bold),
                    ]]
                    for t in module_tasks:
                        indent = "&nbsp;&nbsp;&nbsp;&nbsp;" if norm_text(t.parent_task_id) else ""
                        task_rows.append([
                            Paragraph(f"{indent}{t.task_name}", normal_style),
                            Paragraph(t.department or "-", normal_style),
                            Paragraph(f"{float(t.estimated_hours or 0.0):.2f}", normal_style),
                            Paragraph(t.status or "-", normal_style),
                            Paragraph(t.notes or "-", normal_style),
                        ])

                    table = Table(
                        task_rows,
                        colWidths=[75 * mm, 28 * mm, 15 * mm, 22 * mm, 45 * mm],
                        repeatRows=1,
                    )
                    table.setStyle(self._table_style())
                    elements.append(table)
                else:
                    elements.append(Paragraph("No tasks linked to this module.", normal_style))

                elements.append(Spacer(1, 8))
        else:
            elements.append(Paragraph("No module detail available.", normal_style))

        elements.append(Spacer(1, 10))
        elements.append(Paragraph("Instruction Manuals / Product Documents", section_style))
        if docs:
            rows = [[
                Paragraph("<b>Section</b>", small_bold),
                Paragraph("<b>Document</b>", small_bold),
                Paragraph("<b>Type</b>", small_bold),
                Paragraph("<b>Path</b>", small_bold),
            ]]
            for d in docs:
                rows.append([
                    Paragraph(d.section_name or "-", normal_style),
                    Paragraph(d.doc_name or "-", normal_style),
                    Paragraph(d.doc_type or "-", normal_style),
                    Paragraph(d.file_path or "-", normal_style),
                ])

            table = Table(
                rows,
                colWidths=[25 * mm, 40 * mm, 28 * mm, 85 * mm],
                repeatRows=1,
            )
            table.setStyle(self._table_style())
            elements.append(table)
        else:
            elements.append(Paragraph("No product documents found.", normal_style))

        elements.append(Spacer(1, 10))
        elements.append(Paragraph("Product Work Orders", section_style))
        if workorders:
            rows = [[
                Paragraph("<b>Name</b>", small_bold),
                Paragraph("<b>Stage</b>", small_bold),
                Paragraph("<b>Owner</b>", small_bold),
                Paragraph("<b>Due</b>", small_bold),
                Paragraph("<b>Status</b>", small_bold),
                Paragraph("<b>Notes</b>", small_bold),
            ]]
            for w in workorders:
                rows.append([
                    Paragraph(w.workorder_name or "-", normal_style),
                    Paragraph(w.stage or "-", normal_style),
                    Paragraph(w.owner or "-", normal_style),
                    Paragraph(w.due_date or "-", normal_style),
                    Paragraph(w.status or "-", normal_style),
                    Paragraph(w.notes or "-", normal_style),
                ])

            table = Table(
                rows,
                colWidths=[35 * mm, 25 * mm, 28 * mm, 20 * mm, 22 * mm, 50 * mm],
                repeatRows=1,
            )
            table.setStyle(self._table_style())
            elements.append(table)
        else:
            elements.append(Paragraph("No work orders found for this product.", normal_style))

        doc.build(elements)

    def build_project_report_pdf(self, project_code: str, pdf_path: str) -> None:
        ensure_reportlab_loaded()
        bundle = self.services.projects.get_project_bundle(project_code)
        if not bundle.project:
            raise ValueError("Project not found.")

        project = bundle.project
        module_links = bundle.module_links or []
        project_tasks = bundle.project_tasks or []
        project_docs = bundle.project_documents or []
        workorders = bundle.workorders or []

        project_task_lookup = {t.project_task_id: t for t in project_tasks}
        rollup = {"parts_cost": 0.0, "labour_hours": 0.0, "assembly_quotes": [], "parts": []}
        try:
            rollup = self.services.projects.get_project_rollup(project_code)
        except Exception:
            pass
        blockers = []
        try:
            blockers = self.services.scheduler.get_open_blockers_for_project(project_code)
        except Exception:
            blockers = []

        doc = self._doc(pdf_path)
        title_style, section_style, normal_style, small_bold = self._styles()
        elements = []

        elements.append(Paragraph("Liquimech Live Order Report", title_style))
        elements.append(Spacer(1, 4))

        summary_data = [
            ["Order Code", project.project_code],
            ["Quote Ref", project.quote_ref],
            ["Order Name", project.project_name],
            ["Client Name", project.client_name],
            ["Location", project.location],
            ["Description", project.description],
            ["Linked Product", project.linked_product_code or "-"],
            ["Status", project.status],
            ["Start Date", project.start_date or "-"],
            ["Due Date", project.due_date or "-"],
            ["Assemblies", str(len(module_links))],
            ["Tasks", str(len(project_tasks))],
            ["Documents", str(len(project_docs))],
            ["Work Orders", str(len(workorders))],
            ["Labour Hours", f"{float(bundle.total_hours or 0.0):.2f}"],
            ["Rolled-up Parts Cost", f"{float(rollup.get('parts_cost', 0.0)):.2f}"],
            ["Open Blockers", str(len(blockers))],
        ]
        elements.append(self._summary_table(summary_data))
        elements.append(Spacer(1, 10))

        elements.append(Paragraph("Assembly Quote Rollup", section_style))
        if rollup.get("assembly_quotes"):
            rows = [[
                Paragraph("<b>Assembly</b>", small_bold),
                Paragraph("<b>Qty</b>", small_bold),
                Paragraph("<b>Parts Cost</b>", small_bold),
                Paragraph("<b>Labour Hours</b>", small_bold),
            ]]
            for aq in rollup.get("assembly_quotes", []):
                rows.append([
                    Paragraph(f"{aq.get('assembly_code')} - {aq.get('assembly_name') or ''}", normal_style),
                    Paragraph(str(aq.get('qty', 0)), normal_style),
                    Paragraph(f"{float(aq.get('parts_cost_total', 0.0)):.2f}", normal_style),
                    Paragraph(f"{float(aq.get('hours_total', 0.0)):.2f}", normal_style),
                ])
            table = Table(rows, colWidths=[80 * mm, 18 * mm, 35 * mm, 35 * mm], repeatRows=1)
            table.setStyle(self._table_style())
            elements.append(table)
        else:
            elements.append(Paragraph("No assembly rollup data found.", normal_style))

        elements.append(Spacer(1, 10))

        elements.append(Paragraph("Assembly Execution Status", section_style))
        if module_links:
            rows = [[
                Paragraph("<b>Order</b>", small_bold),
                Paragraph("<b>Assembly Code</b>", small_bold),
                Paragraph("<b>Source</b>", small_bold),
                Paragraph("<b>Qty</b>", small_bold),
                Paragraph("<b>Stage</b>", small_bold),
                Paragraph("<b>Status</b>", small_bold),
                Paragraph("<b>Dependency</b>", small_bold),
                Paragraph("<b>Notes</b>", small_bold),
            ]]
            for m in module_links:
                rows.append([
                    Paragraph(str(m.module_order), normal_style),
                    Paragraph(m.module_code or "-", normal_style),
                    Paragraph(f"{m.source_type} / {m.source_code}", normal_style),
                    Paragraph(str(m.module_qty), normal_style),
                    Paragraph(m.stage or "-", normal_style),
                    Paragraph(m.status or "-", normal_style),
                    Paragraph(m.dependency_module_code or "-", normal_style),
                    Paragraph(m.notes or "-", normal_style),
                ])

            table = Table(
                rows,
                colWidths=[12 * mm, 28 * mm, 35 * mm, 12 * mm, 24 * mm, 22 * mm, 28 * mm, 45 * mm],
                repeatRows=1,
            )
            table.setStyle(self._table_style())
            elements.append(table)
        else:
            elements.append(Paragraph("No project modules found.", normal_style))

        elements.append(Spacer(1, 10))
        elements.append(Paragraph("Project Task Execution", section_style))
        if project_tasks:
            rows = [[
                Paragraph("<b>Module</b>", small_bold),
                Paragraph("<b>Task</b>", small_bold),
                Paragraph("<b>Dept</b>", small_bold),
                Paragraph("<b>Hours</b>", small_bold),
                Paragraph("<b>Stage</b>", small_bold),
                Paragraph("<b>Status</b>", small_bold),
                Paragraph("<b>Assigned</b>", small_bold),
                Paragraph("<b>Dependency</b>", small_bold),
            ]]
            for t in project_tasks:
                dep_display = ""
                if norm_text(t.dependency_task_id) and t.dependency_task_id in project_task_lookup:
                    dep = project_task_lookup[t.dependency_task_id]
                    dep_display = f"{dep.module_code} | {dep.task_name}"
                elif norm_text(t.dependency_task_id):
                    dep_display = t.dependency_task_id

                rows.append([
                    Paragraph(t.module_code or "-", normal_style),
                    Paragraph(t.task_name or "-", normal_style),
                    Paragraph(t.department or "-", normal_style),
                    Paragraph(f"{float(t.estimated_hours or 0.0):.2f}", normal_style),
                    Paragraph(t.stage or "-", normal_style),
                    Paragraph(t.status or "-", normal_style),
                    Paragraph(t.assigned_to or "-", normal_style),
                    Paragraph(dep_display or "-", normal_style),
                ])

            table = Table(
                rows,
                colWidths=[24 * mm, 42 * mm, 22 * mm, 12 * mm, 22 * mm, 22 * mm, 25 * mm, 35 * mm],
                repeatRows=1,
            )
            table.setStyle(self._table_style())
            elements.append(table)
        else:
            elements.append(Paragraph("No project tasks found.", normal_style))

        elements.append(Spacer(1, 10))
        elements.append(Paragraph("Project Documents", section_style))
        if project_docs:
            rows = [[
                Paragraph("<b>Section</b>", small_bold),
                Paragraph("<b>Document</b>", small_bold),
                Paragraph("<b>Type</b>", small_bold),
                Paragraph("<b>Path</b>", small_bold),
            ]]
            for d in project_docs:
                rows.append([
                    Paragraph(d.section_name or "-", normal_style),
                    Paragraph(d.doc_name or "-", normal_style),
                    Paragraph(d.doc_type or "-", normal_style),
                    Paragraph(d.file_path or "-", normal_style),
                ])

            table = Table(
                rows,
                colWidths=[25 * mm, 40 * mm, 28 * mm, 85 * mm],
                repeatRows=1,
            )
            table.setStyle(self._table_style())
            elements.append(table)
        else:
            elements.append(Paragraph("No project documents found.", normal_style))

        elements.append(Spacer(1, 10))
        elements.append(Paragraph("Project Work Orders", section_style))
        if workorders:
            rows = [[
                Paragraph("<b>Name</b>", small_bold),
                Paragraph("<b>Stage</b>", small_bold),
                Paragraph("<b>Owner</b>", small_bold),
                Paragraph("<b>Due</b>", small_bold),
                Paragraph("<b>Status</b>", small_bold),
                Paragraph("<b>Notes</b>", small_bold),
            ]]
            for w in workorders:
                rows.append([
                    Paragraph(w.workorder_name or "-", normal_style),
                    Paragraph(w.stage or "-", normal_style),
                    Paragraph(w.owner or "-", normal_style),
                    Paragraph(w.due_date or "-", normal_style),
                    Paragraph(w.status or "-", normal_style),
                    Paragraph(w.notes or "-", normal_style),
                ])

            table = Table(
                rows,
                colWidths=[35 * mm, 25 * mm, 28 * mm, 20 * mm, 22 * mm, 50 * mm],
                repeatRows=1,
            )
            table.setStyle(self._table_style())
            elements.append(table)
        else:
            elements.append(Paragraph("No work orders found for this project.", normal_style))

        elements.append(Spacer(1, 10))
        elements.append(Paragraph("Open Blockers", section_style))
        if blockers:
            rows = [[
                Paragraph("<b>Type</b>", small_bold),
                Paragraph("<b>Item</b>", small_bold),
                Paragraph("<b>Depends On</b>", small_bold),
                Paragraph("<b>Current Status</b>", small_bold),
                Paragraph("<b>Dependency Status</b>", small_bold),
            ]]
            for b in blockers:
                item = b.get("task_name", "") if b.get("type") == "TASK" else b.get("module_code", "")
                rows.append([
                    Paragraph(str(b.get("type", "-")), normal_style),
                    Paragraph(str(item or "-"), normal_style),
                    Paragraph(str(b.get("depends_on", "-")), normal_style),
                    Paragraph(str(b.get("current_status", "-")), normal_style),
                    Paragraph(str(b.get("dependency_status", "-")), normal_style),
                ])

            table = Table(
                rows,
                colWidths=[16 * mm, 45 * mm, 45 * mm, 35 * mm, 35 * mm],
                repeatRows=1,
            )
            table.setStyle(self._table_style())
            elements.append(table)
        else:
            elements.append(Paragraph("No open blockers found.", normal_style))

        doc.build(elements)

    # ========================================================
    # Styling helpers
    # ========================================================

    def _doc(self, pdf_path: str) -> SimpleDocTemplate:
        return SimpleDocTemplate(
            pdf_path,
            pagesize=A4,
            leftMargin=AppConfig.PDF_MARGIN_MM * mm,
            rightMargin=AppConfig.PDF_MARGIN_MM * mm,
            topMargin=AppConfig.PDF_MARGIN_MM * mm,
            bottomMargin=AppConfig.PDF_MARGIN_MM * mm,
        )

    def _styles(self):
        styles = getSampleStyleSheet()

        title_style = ParagraphStyle(
            "TitleStyle",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=18,
            textColor=colors.HexColor(AppConfig.COLOR_PRIMARY),
            spaceAfter=10,
        )

        section_style = ParagraphStyle(
            "SectionStyle",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=12,
            textColor=colors.HexColor(AppConfig.COLOR_PRIMARY),
            spaceAfter=6,
            spaceBefore=8,
        )

        normal_style = ParagraphStyle(
            "NormalClean",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=8.7,
            leading=11,
            textColor=colors.black,
        )

        small_bold = ParagraphStyle(
            "SmallBold",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=8.5,
            textColor=colors.black,
        )

        return title_style, section_style, normal_style, small_bold

    def _table_style(self) -> TableStyle:
        return TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(AppConfig.COLOR_ACCENT)),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8.2),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor(AppConfig.COLOR_GRID)),
            ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#8095AA")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7FAFD")]),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ])

    def _summary_table(self, summary_data: List[List[str]]) -> Table:
        table = Table(summary_data, colWidths=[45 * mm, 125 * mm])
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor(AppConfig.COLOR_ACCENT_SOFT)),
            ("TEXTCOLOR", (0, 0), (-1, -1), colors.black),
            ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ("FONTNAME", (1, 0), (1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 8.8),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor(AppConfig.COLOR_GRID)),
            ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#8095AA")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, colors.HexColor("#F8FBFE")]),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        return table
