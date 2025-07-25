import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4, landscape
from reportlab.platypus import (
    SimpleDocTemplate, LongTable, TableStyle, Paragraph, Spacer, PageBreak, Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import json
import random                      # --- MODIFICADO: para sortear paletas ---

# --------------------------------------------------------------------------
# --- MODIFICADO: 5 PALETAS COMPLETAS --------------------------------------
# Cada diccionario define los 5 “roles” de color que usa el PDF
# --------------------------------------------------------------------------
PALETTES = [
    # Paleta original
    {
        "HEADER_BG": "#2D2A32",
        "ACCENT":    "#DDD92A",
        "GRID":      "#EAE151",
        "ROW_ODD":   "#EEEFA8",
        "ROW_EVEN":  "#FAFDF6",
    },
    # Paleta 1  (540d6e…)
    {
        "HEADER_BG": "#3590f3",
        "ACCENT":    "#8FB8ED",
        "GRID":      "#62BFED",
        "ROW_ODD":   "#F1E3F3",
        "ROW_EVEN":  "#C2BBF0",
    },
    # Paleta 2  (363537…)
    {
        "HEADER_BG": "#080357",
        "ACCENT":    "#ff9f1c",
        "GRID":      "#ffc15e",
        "ROW_ODD":   "#d6ffb7",
        "ROW_EVEN":  "#f5ff90",
    },
    # Paleta 3  (094074…)
    {
        "HEADER_BG": "#373737",
        "ACCENT":    "#FCE694",
        "GRID":      "#F6FEAA",
        "ROW_ODD":   "#C7DFC5",
        "ROW_EVEN":  "#C1DBE3",
    },
    # Paleta 4  (e0acd5…)
    {
        "HEADER_BG": "#6a3e37",
        "ACCENT":    "#3993dd",
        "GRID":      "#29e7cd",
        "ROW_ODD":   "#e0acd5",
        "ROW_EVEN":  "#f4ebe8",
    },
]

# Elegir una paleta y actualizar los colores globales
_palette = None  # almacena la última paleta seleccionada

def choose_palette():
    """Escoge una paleta al azar y actualiza las variables de color."""
    global COLOR_1, COLOR_2, COLOR_3, COLOR_4, COLOR_5, _palette
    _palette = random.choice(PALETTES)
    COLOR_1 = colors.HexColor(_palette["HEADER_BG"])
    COLOR_2 = colors.HexColor(_palette["ACCENT"])
    COLOR_3 = colors.HexColor(_palette["GRID"])
    COLOR_4 = colors.HexColor(_palette["ROW_ODD"])
    COLOR_5 = colors.HexColor(_palette["ROW_EVEN"])

# Paleta inicial para la interfaz
choose_palette()
# --------------------------------------------------------------------------

# Archivo donde se guardan los anchos de columna utilizados
CONFIG_FILE = "column_widths.json"

# Ruta de la imagen utilizada como marca de agua. Modifica esta
# variable para apuntar a la ubicación de tu archivo PNG
WATERMARK_IMAGE = r"C:\Users\El Pela Flow\OneDrive\Documentos\ReportLab plantilla leads\watermark.png"

# --- Encabezado y pie ---
def add_page_elements(canvas, doc):
    """Dibuja en cada página el encabezado, el pie y la marca de agua."""

    # --- Marca de agua ---
    if WATERMARK_IMAGE and os.path.exists(WATERMARK_IMAGE):
        canvas.saveState()
        canvas.setFillAlpha(0.1)
        width, height = doc.pagesize
        canvas.drawImage(
            WATERMARK_IMAGE,
            0, 0, width=width, height=height,
            preserveAspectRatio=True, mask='auto'
        )
        canvas.restoreState()

    canvas.saveState()
    canvas.setFont("Helvetica-Bold", 10)
    canvas.setFillColor(COLOR_1)
    canvas.drawString(4*inch, letter[1] - 0.7*inch, "Reporte de Leads - Mi Empresa")
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(COLOR_2)
    page_num = canvas.getPageNumber()
    canvas.drawString(0.5*inch, 0.5*inch, f"Página {page_num}")
    canvas.drawRightString(
        letter[0] - 0.5*inch,
        0.5*inch,
        f"Generado el {datetime.now().strftime('%d/%m/%Y')}"
    )
    canvas.setStrokeColor(COLOR_2)
    canvas.setLineWidth(1)
    canvas.line(
        0.5*inch, letter[1] - 1*inch,
        letter[0] - 0.5*inch, letter[1] - 1*inch
    )
    canvas.restoreState()


# --- Generación del PDF ---
def generate_pdf(csv_file, col_widths_px=None, glossary_pdf=None):
    """Genera un reporte PDF a partir de un archivo CSV."""
    # Escoger una paleta distinta en cada ejecución
    choose_palette()
    try:
        # Leer CSV
        df = pd.read_csv(csv_file, sep=';', encoding='utf-8', engine='python')

        # Calcular anchos proporcionales
        total_width = landscape(letter)[0] - (0.5 + 0.5) * inch
        min_w = 0.6 * inch

        if col_widths_px:
            px_total = sum(col_widths_px)
            col_widths = [w / px_total * total_width for w in col_widths_px]
        else:
            max_chars = [
                max(df[col].astype(str).map(len).max(), len(col))
                for col in df.columns
            ]
            sum_chars = sum(max_chars)
            raw_w = [mc / sum_chars * total_width for mc in max_chars]
            col_widths = [max(w, min_w) for w in raw_w]

        # Estilos de párrafo
        styles = getSampleStyleSheet()

        header_style = ParagraphStyle(
            name="Header",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=5,
            leading=8,
            alignment=1,
            textColor=COLOR_5     # Texto claro sobre fondo oscuro
        )

        cell_style = ParagraphStyle(
            name="Cell",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=5,
            leading=6,
            alignment=1
        )

        # Preparar data
        data = [[Paragraph(col, header_style) for col in df.columns]]
        for row in df.itertuples(index=False):
            data.append([Paragraph(str(c) if pd.notna(c) else "", cell_style)
                         for c in row])

        # Crear y poblar el PDF
        csv_dir = os.path.dirname(csv_file)
        stylized_pdf = os.path.join(csv_dir, "reporte_leads_estilizado.pdf")
        doc = SimpleDocTemplate(
            stylized_pdf,
            pagesize=landscape(A4),
            leftMargin=0.5*inch, rightMargin=0.5*inch,
            topMargin=1*inch, bottomMargin=0.7*inch
        )

        elements = []

        # --- Encabezado de tabla ---
        elements.append(Paragraph("Listado de contactos de negocios", styles["Title"]))
        elements.append(Spacer(1, 0.1 * inch))
        elements.append(Paragraph(
            f"Generado automáticamente el {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            styles["Normal"]
        ))
        elements.append(Spacer(1, 0.2 * inch))

        table = LongTable(data, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), COLOR_1),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, 0), 6),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('VALIGN',   (0, 1), (-1, -1), 'MIDDLE'),
            ('BACKGROUND', (0, 1), (-1, -1), COLOR_4),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [COLOR_4, COLOR_5]),
            ('GRID', (0, 0), (-1, -1), 0.25, COLOR_3),
        ]))

        elements.append(table)
        doc.build(elements, onFirstPage=add_page_elements, onLaterPages=add_page_elements)

        if glossary_pdf and os.path.exists(glossary_pdf):
            from PyPDF2 import PdfMerger

            merger = PdfMerger()
            merger.append(glossary_pdf)
            merger.append(stylized_pdf)

            final_pdf = os.path.join(csv_dir, "reporte_completo.pdf")
            with open(final_pdf, "wb") as fout:
                merger.write(fout)
            merger.close()

            # Eliminar el PDF intermedio para entregar un único reporte
            try:
                os.remove(stylized_pdf)
            except Exception:
                pass

            messagebox.showinfo("Éxito", f"PDF combinado generado:\n{final_pdf}")
        else:
            messagebox.showinfo("Éxito", f"PDF generado:\n{stylized_pdf}")

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo generar el PDF:\n{e}")

# --- Interfaz gráfica ---
class PDFGeneratorApp:
    """Interfaz gráfica para seleccionar un CSV y generar el PDF."""
    def __init__(self, root):
        self.root = root
        root.title("Generador de PDF desde CSV")
        root.geometry("450x180")

        tk.Label(root, text="Selecciona un CSV de leads", font=("Helvetica", 12)).pack(pady=10)
        tk.Button(
            root,
            text="Seleccionar archivo...",
            command=self.select_csv,
            bg=_palette["ACCENT"],           # --- MODIFICADO: usa color de la paleta
            fg=_palette["HEADER_BG"]
        ).pack()
        tk.Button(
            root,
            text="Seleccionar Glosario (PDF)",
            command=self.select_glossary_pdf
        ).pack(pady=2)
        self.file_label = tk.Label(root, text="Ningún archivo seleccionado", font=("Helvetica", 10))
        self.file_label.pack(pady=8)
        tk.Button(
            root,
            text="Generar PDF",
            command=self.on_generate,
            bg=_palette["HEADER_BG"],       # --- MODIFICADO: usa color de la paleta
            fg=_palette["ROW_EVEN"]
        ).pack(pady=5)

        self.csv_file = None
        self.glossary_pdf = None
        self.df = None
        self.sort_states = {}

    def select_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if path:
            self.csv_file = path
            self.file_label.config(text=os.path.basename(path))
            try:
                self.df = pd.read_csv(path, sep=';', encoding='utf-8', engine='python')
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo leer el CSV:\n{e}")
                self.csv_file = None
                self.df = None

    def select_glossary_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.glossary_pdf = path

    def on_generate(self):
        if not self.csv_file:
            messagebox.showwarning("Atención", "Debes seleccionar un archivo CSV primero.")
        else:
            self.open_adjust_window()

    def open_adjust_window(self):
        if self.df is None:
            return
        win = tk.Toplevel(self.root)
        win.title("Ajustar anchos de columnas")
        tree = ttk.Treeview(win, columns=list(self.df.columns), show='headings')

        # Inicializa estados de ordenamiento
        self.sort_states = {c: False for c in self.df.columns}

        def sort_column(col):
            reverse = self.sort_states.get(col, False)
            self.df.sort_values(by=col, ascending=not reverse, inplace=True)
            self.sort_states[col] = not reverse

            for c in self.df.columns:
                indicator = ''
                if c == col:
                    indicator = ' \u25B2' if not reverse else ' \u25BC'
                tree.heading(c, text=c + indicator, command=lambda col=c: sort_column(col))

            for row in tree.get_children():
                tree.delete(row)
            for row in self.df.head(5).itertuples(index=False):
                tree.insert('', 'end', values=list(row))

        # Cargar anchos guardados
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    saved_widths = json.load(f)
            except Exception:
                saved_widths = {}
        else:
            saved_widths = {}

        for col in self.df.columns:
            tree.heading(col, text=col, command=lambda c=col: sort_column(c))
            width = saved_widths.get(col, 100)
            tree.column(col, width=width, stretch=True)
        for row in self.df.head(5).itertuples(index=False):
            tree.insert('', 'end', values=list(row))
        tree.pack(fill='both', expand=True, padx=10, pady=10)

        def confirm():
            widths = [tree.column(c)['width'] for c in self.df.columns]
            try:
                with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                    json.dump(dict(zip(self.df.columns, widths)), f)
            except Exception:
                pass
            win.destroy()
            generate_pdf(self.csv_file, widths, glossary_pdf=self.glossary_pdf)

        ttk.Button(win, text="Generar PDF", command=confirm).pack(pady=5)

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFGeneratorApp(root)
    root.mainloop()
