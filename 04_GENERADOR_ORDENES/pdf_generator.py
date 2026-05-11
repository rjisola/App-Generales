from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, Flowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
import os

class RoundedTableWrapper(Flowable):
    """Clase auxiliar compatible con ReportLab para dibujar bordes redondeados detrás de una tabla"""
    def __init__(self, table, width, height, radius=10):
        Flowable.__init__(self)
        self.table = table
        self.width = width
        self.height = height
        self.radius = radius

    def draw(self):
        canvas = self.canv
        canvas.saveState()
        canvas.setStrokeColor(colors.black)
        canvas.setLineWidth(1)
        # Dibujar el rectángulo redondeado
        canvas.roundRect(0, 0, self.width, self.height, self.radius, stroke=1, fill=0)
        canvas.restoreState()
        # Dibujar la tabla encima del fondo redondeado
        self.table.drawOn(canvas, 0, 0)

    def wrap(self, availWidth, availHeight):
        self.table.wrap(self.width, self.height)
        return self.width, self.height

class GeneradorOrdenPDF:
    def __init__(self, output_path):
        self.output_path = output_path
        self.styles = getSampleStyleSheet()

    def format_n(self, value):
        try:
            val = float(value)
            formatted = f"{val:,.2f}"
            return formatted.replace(",", "X").replace(".", ",").replace("X", ".")
        except (ValueError, TypeError):
            return "0,00"

    def generar(self, orden_data, items):
        doc = SimpleDocTemplate(
            self.output_path, 
            pagesize=A4, 
            rightMargin=1*cm, 
            leftMargin=1*cm, 
            topMargin=1*cm, 
            bottomMargin=1*cm
        )
        elements = []

        # Estilos
        bold_style = ParagraphStyle('BoldStyle', parent=self.styles['Normal'], fontSize=9, fontName='Helvetica-Bold')
        normal_style = ParagraphStyle('NormalStyle', parent=self.styles['Normal'], fontSize=9)
        small_style = ParagraphStyle('SmallStyle', parent=self.styles['Normal'], fontSize=8)
        header_title = ParagraphStyle('HeaderTitle', parent=self.styles['Normal'], fontSize=14, fontName='Helvetica-Bold')

        # --- SECCION 1: ENCABEZADO ---
        logo_path = os.path.join(os.path.dirname(__file__), "logo_empresa.png")
        if os.path.exists(logo_path) and os.path.getsize(logo_path) > 1000:
            img = Image(logo_path, width=4*cm, height=2*cm)
            header_col1 = img
        else:
            header_col1 = Paragraph("<br/><br/><font size=18>CARJOR</font>", header_title)

        header_data = [
            [header_col1, "", Paragraph(f"O.Compra Nro: {orden_data['numero_orden']}", header_title)],
            [Paragraph("Domicilio: 3 de Febrero 59", small_style), "", Paragraph(f"Fecha: {orden_data['fecha']}", bold_style)],
            [Paragraph("Zárate (2800) - Buenos Aires", small_style), "", Paragraph("F. Expiración: 20/12/2030", small_style)],
            [Paragraph("Teléfono: (03487) 434-739 | CUIT: 30-70921165-6", small_style), "", ""],
            [Paragraph("Email: empresacarjor@gmail.com | Ing. Brutos: 30-70921165-6", small_style), "", ""],
            [Paragraph("Iva Responsable Inscripto | Inicio act.: Julio de 2005", small_style), "", ""]
        ]
        
        t_head = Table(header_data, colWidths=[10*cm, 2*cm, 7*cm])
        t_head.setStyle(TableStyle([
            ('SPAN', (0,0), (1,0)),
            ('ALIGN', (2,0), (2,2), 'RIGHT'),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ]))
        elements.append(t_head)
        elements.append(Spacer(1, 0.5*cm))

        # --- SECCION 2: DATOS DEL PROVEEDOR ---
        # Incluimos el nombre del proveedor DENTRO de la tabla para que todo quede en el mismo cuadro redondeado
        prov_data = [
            [Paragraph(f"<b>Proveedor:</b> {orden_data['proveedor_nombre']}", normal_style), ""],
            [Paragraph(f"<b>Dirección:</b> {orden_data.get('domicilio', '')}", normal_style), Paragraph("Enviar FACTURA a facturascarjor@gmail.com", small_style)],
            [Paragraph(f"<b>I.V.A.:</b> {orden_data.get('categoria_iva', '')}", normal_style), Paragraph(f"<b>CUIT:</b> {orden_data.get('cuit', '')}", normal_style)]
        ]
        t_prov = Table(prov_data, colWidths=[10*cm, 9*cm])
        t_prov.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
            ('RIGHTPADDING', (0,0), (-1,-1), 10),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            # NO USAMOS 'BOX' ni 'GRID' para evitar esquinas dobles
        ]))
        
        elements.append(RoundedTableWrapper(t_prov, 19*cm, 2.3*cm, radius=8))
        elements.append(Spacer(1, 0.5*cm))

        # --- SECCION 3: TABLA DE ITEMS ---
        data_items = [["Item", "Descripción", "Cant.", "Precio", "Importe"]]
        for i, item in enumerate(items, 1):
            data_items.append([
                str(i),
                item['descripcion'],
                self.format_n(item['cantidad']),
                self.format_n(item['precio_unitario']),
                self.format_n(item['total_item'])
            ])
        
        while len(data_items) < 21:
            data_items.append(["", "", "", "", ""])

        row_heights = [0.5*cm] * len(data_items)
        t_items = Table(data_items, colWidths=[1.0*cm, 9.0*cm, 2.0*cm, 3.5*cm, 3.5*cm], rowHeights=row_heights, repeatRows=1)
        t_items.setStyle(TableStyle([
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            # Solo GRID interno, NO BOX externo
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('ALIGN', (2,0), (-1,-1), 'CENTER'),
            ('ALIGN', (1,1), (1,-1), 'LEFT'),
            ('ALIGN', (3,1), (4,-1), 'RIGHT'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        
        elements.append(RoundedTableWrapper(t_items, 19*cm, len(data_items)*0.5*cm, radius=8))
        elements.append(Spacer(1, 0.5*cm))

        # --- SECCION 4: PIE DE PAGINA Y TOTALES ---
        foot_data = [
            [Paragraph(f"<b>AUTORIZA:</b> {orden_data.get('autorizado', '')}", small_style), "SUB TOTAL", f"$ {self.format_n(orden_data['subtotal'])}"],
            [Paragraph(f"<b>FECHA ENT.:</b> {orden_data.get('fecha_entrega', '')}", small_style), f"PERCEP IIBB ({orden_data.get('p_iibb', 0)}%)", f"$ {self.format_n(orden_data.get('iibb', 0))}"],
            [Paragraph(f"<b>RETIRA:</b> {orden_data.get('retira', '')}", small_style), f"LEY 23.966 ({orden_data.get('p_l23', 0)}%)", f"$ {self.format_n(orden_data.get('ley23966', 0))}"],
            [Paragraph(f"<b>DESTINO:</b> {orden_data.get('destino', '')}", small_style), f"LEY 27.430 ({orden_data.get('p_l27', 0)}%)", f"$ {self.format_n(orden_data.get('ley27430', 0))}"],
            [Paragraph(f"<b>OBRA:</b> {orden_data.get('obra', '')}", small_style), f"I.V.A. ({orden_data.get('p_iva', 0)}%)", f"$ {self.format_n(orden_data['iva'])}"],
            [Paragraph(f"<b>F. DE PAGO:</b> {orden_data.get('forma_pago', '')}", small_style), "TOTAL", f"$ {self.format_n(orden_data['total'])}"]
        ]
        
        t_foot = Table(foot_data, colWidths=[12*cm, 3.5*cm, 3.5*cm], rowHeights=[0.5*cm]*6)
        t_foot.setStyle(TableStyle([
            ('ALIGN', (1,0), (1,-1), 'RIGHT'),
            ('ALIGN', (2,0), (2,-1), 'RIGHT'),
            ('FONTNAME', (1,5), (2,5), 'Helvetica-Bold'),
            ('FONTSIZE', (1,5), (2,5), 10),
            # GRID interno SOLO para la parte de totales (columnas 1 y 2)
            ('GRID', (1,0), (2,5), 0.5, colors.black),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('LEFTPADDING', (0,0), (0,-1), 10),
        ]))
        
        elements.append(RoundedTableWrapper(t_foot, 19*cm, 6*0.5*cm, radius=8))

        doc.build(elements)
        return self.output_path
