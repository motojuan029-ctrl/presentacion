from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# ================= CONFIGURACIÓN ESTÉTICA =================
FONDO_NEGRO = RGBColor(0x12, 0x12, 0x12)   # #121212
DORADO = RGBColor(0xEA, 0xB3, 0x08)        # #EAB308
BLANCO = RGBColor(0xFF, 0xFF, 0xFF)
GRIS_CLARO = RGBColor(0xCC, 0xCC, 0xCC)

def set_slide_background(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_title(slide, text, left=Inches(0.5), top=Inches(0.3), width=Inches(9), height=Inches(1)):
    tb = slide.shapes.add_textbox(left, top, width, height)
    p = tb.text_frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = DORADO
    p.alignment = PP_ALIGN.LEFT

def add_subtitle(slide, text, left=Inches(0.5), top=Inches(1.2), width=Inches(9), height=Inches(0.5)):
    tb = slide.shapes.add_textbox(left, top, width, height)
    p = tb.text_frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(20)
    p.font.color.rgb = GRIS_CLARO

def add_bullet(slide, text, left=Inches(0.7), top=Inches(2), width=Inches(8.5), height=Inches(0.4), level=0):
    tb = slide.shapes.add_textbox(left + Inches(level*0.3), top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(18) if level == 0 else Pt(16)
    p.font.color.rgb = BLANCO if level == 0 else GRIS_CLARO
    p.level = level

def add_centered_title(slide, text):
    tb = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1))
    p = tb.text_frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = DORADO
    p.alignment = PP_ALIGN.CENTER

def add_quote(slide, text, top=Inches(3)):
    tb = slide.shapes.add_textbox(Inches(1), top, Inches(8), Inches(1.2))
    p = tb.text_frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(28)
    p.font.italic = True
    p.font.color.rgb = GRIS_CLARO
    p.alignment = PP_ALIGN.CENTER

# Crear presentación
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

# ================= DIA 1: PORTADA =================
slide1 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide1, FONDO_NEGRO)
add_centered_title(slide1, "UNITURNO")
sub = slide1.shapes.add_textbox(Inches(2), Inches(3.5), Inches(6), Inches(0.8))
sub.text_frame.text = "El valor del tiempo, optimizado por ingeniería"
sub.text_frame.paragraphs[0].font.size = Pt(24)
sub.text_frame.paragraphs[0].font.color.rgb = GRIS_CLARO
sub.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
integrantes = slide1.shapes.add_textbox(Inches(3), Inches(5.5), Inches(4), Inches(1))
integrantes.text_frame.text = "Juan Brito · José Hernández\nJoaco Torregroza · Rafael Puertas\nJary Villar"
integrantes.text_frame.paragraphs[0].font.size = Pt(14)
integrantes.text_frame.paragraphs[0].font.color.rgb = BLANCO
integrantes.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

# ================= DIA 2: PROBLEMA GLOBAL =================
slide2 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide2, FONDO_NEGRO)
add_title(slide2, "El problema global")
add_subtitle(slide2, "El tiempo no se recupera")
add_bullet(slide2, "• Infraestructuras físicas obsoletas generan aglomeraciones masivas")
add_bullet(slide2, "• Afecta universidades, empresas y comedores industriales")
add_bullet(slide2, "• Falta de control predictivo de la demanda")
add_bullet(slide2, "• Pérdida sistemática de productividad y bienestar")

# ================= DIA 3: IMPACTO INVISIBLE =================
slide3 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide3, FONDO_NEGRO)
add_title(slide3, "El Impacto Invisible")
add_subtitle(slide3, "Datos del piloto en Universidad del Magdalena")
add_bullet(slide3, "• 22,000 estudiantes afectados diariamente")
add_bullet(slide3, "• 1 hora promedio de espera por estudiante")
add_bullet(slide3, "• Colapso total en horas pico (12:00 – 14:00)")
add_bullet(slide3, "• 45-60 minutos perdidos que podrían ser de estudio o descanso")

# ================= DIA 4: CONTEXTO PILOTO =================
slide4 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide4, FONDO_NEGRO)
add_title(slide4, "Contexto Piloto")
add_subtitle(slide4, "Comedor universitario – Unimagdalena")
add_bullet(slide4, "• 22,000 estudiantes con infraestructura limitada")
add_bullet(slide4, "• Sistema actual: filas tradicionales → cuellos de botella")
add_bullet(slide4, "• Horario crítico: 11:30 am – 1:30 pm")
add_bullet(slide4, "• Frustración y abandono del servicio por falta de visibilidad")

# ================= DIA 5: EL COLAPSO NO ES EXCLUSIVO =================
slide5 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide5, FONDO_NEGRO)
add_centered_title(slide5, "Pero este colapso")
add_centered_title(slide5, "no es exclusivo")
# Mover el segundo título un poco más abajo
tb2 = slide5.shapes.add_textbox(Inches(1), Inches(3.2), Inches(8), Inches(1))
p2 = tb2.text_frame.paragraphs[0]
p2.text = "no es exclusivo"
p2.font.size = Pt(44)
p2.font.bold = True
p2.font.color.rgb = DORADO
p2.alignment = PP_ALIGN.CENTER

# ================= DIA 6: SOLUCIÓN =================
slide6 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide6, FONDO_NEGRO)
add_centered_title(slide6, "TurnoUni")
add_subtitle(slide6, "Sistema inteligente de gestión de demanda modular")
add_bullet(slide6, "• Erradicar las filas por completo")
add_bullet(slide6, "• Optimizar el flujo humano en cualquier organización")
add_bullet(slide6, "• Estandarizar la asignación de demanda digital")
add_bullet(slide6, "• Puente tecnológico escalable a nivel global")

# ================= DIA 7: REPLICABILIDAD =================
slide7 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide7, FONDO_NEGRO)
add_quote(slide7, "Para un inversor, una solución solo es atractiva si es replicable.", top=Inches(3.2))

# ================= DIA 8: ALTERNATIVAS DE DISEÑO =================
slide8 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide8, FONDO_NEGRO)
add_title(slide8, "Alternativas de Diseño")
headers = ["Solución", "Ventajas", "Desventajas", "Estado"]
rows = [
    ["Tótems físicos", "Presencia visual", "Alto costo, mantenimiento", "Descartado"],
    ["Tickets de papel", "Bajo costo inicial", "Insostenible, sin datos", "Descartado"],
    ["WebApp Cloud", "Cero hardware, escalable", "Requiere conectividad", "✅ Seleccionada"]
]
table = slide8.shapes.add_table(len(rows)+1, len(headers), Inches(0.8), Inches(2.2), Inches(8.4), Inches(2)).table
for i, h in enumerate(headers):
    cell = table.cell(0, i)
    cell.text = h
    cell.text_frame.paragraphs[0].font.bold = True
    cell.text_frame.paragraphs[0].font.color.rgb = DORADO
for i, row in enumerate(rows):
    for j, val in enumerate(row):
        cell = table.cell(i+1, j)
        cell.text = val
        cell.text_frame.paragraphs[0].font.color.rgb = BLANCO

# ================= DIA 9: EVALUACIÓN DE IMPACTO =================
slide9 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide9, FONDO_NEGRO)
add_title(slide9, "Evaluación de Impacto")
add_bullet(slide9, "• Social: Dignifica el tiempo, reduce estrés académico/laboral")
add_bullet(slide9, "• Operativo: Eliminación de filas – aplanamiento del 100% en horas pico")
add_bullet(slide9, "• Ambiental: Cero uso de papel/plástico, organizaciones sostenibles")

# ================= DIA 10: DEMO - AUTENTICACIÓN =================
slide10 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide10, FONDO_NEGRO)
add_title(slide10, "Demo: Autenticación Segura")
add_subtitle(slide10, "Login institucional escalable a corporaciones")
add_bullet(slide10, "• Integración con credenciales universitarias (SSO)")
add_bullet(slide10, "• Preparado para LDAP / OAuth corporativo")
add_bullet(slide10, "• Cifrado de datos y sesiones JWT")

# ================= DIA 11: DEMO - GESTIÓN DE CUPOS =================
slide11 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide11, FONDO_NEGRO)
add_title(slide11, "Gestión de Cupos Dinámicos")
add_subtitle(slide11, "Franjas horarias y control de aforo en tiempo real")
add_bullet(slide11, "• Visualización de cupos disponibles")
add_bullet(slide11, "• Bloqueo automático al alcanzar límite de aforo")
add_bullet(slide11, "• Sugerencia de horarios alternativos")
add_bullet(slide11, "• Cancelación / reagendamiento con un toque")

# ================= DIA 12: DEMO - QR =================
slide12 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide12, FONDO_NEGRO)
add_title(slide12, "Validación por QR")
add_subtitle(slide12, "Llave de acceso rápida (< 2 segundos)")
add_bullet(slide12, "• Generación de QR único y cifrado al reservar")
add_bullet(slide12, "• Escaneo instantáneo – sin biometría lenta")
add_bullet(slide12, "• Flujo continuo y ordenado en el punto de control")

# ================= DIA 13: PANEL ADMINISTRATIVO =================
slide13 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide13, FONDO_NEGRO)
add_title(slide13, "Panel Administrativo")
add_subtitle(slide13, "Monitoreo y control para gerentes")
add_bullet(slide13, "• Dashboard con aforo en tiempo real")
add_bullet(slide13, "• Ajuste de cupos por franja horaria")
add_bullet(slide13, "• Analytics: horarios pico, rotación, desperdicio")
add_bullet(slide13, "• Exportación de datos para optimización operativa")

# ================= DIA 14: ARQUITECTURA TÉCNICA =================
slide14 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide14, FONDO_NEGRO)
add_title(slide14, "Arquitectura Técnica")
add_subtitle(slide14, "Pipeline de desarrollo (5 fases)")
add_bullet(slide14, "1. VS Code + Claude Code CLI + Supabase + Railway")
add_bullet(slide14, "2. Backend: API REST ultrarrápida (PostgreSQL)")
add_bullet(slide14, "3. Frontend: componentes interactivos en la nube")
add_bullet(slide14, "4. Despliegue continuo automatizado en Railway")
add_bullet(slide14, "5. Pruebas en producción con alta carga simulada")

# ================= DIA 15: ESCALABILIDAD =================
slide15 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide15, FONDO_NEGRO)
add_title(slide15, "Escalabilidad y Generalización")
add_subtitle(slide15, "Modelo SaaS para múltiples industrias")
add_bullet(slide15, "• Universidades → comedores, bibliotecas")
add_bullet(slide15, "• Hospitales → control de citas y visitas")
add_bullet(slide15, "• Bancos → gestión de turnos en sucursales")
add_bullet(slide15, "• Comedores corporativos → fila cero para empleados")
add_bullet(slide15, "• Cualquier organización con alta afluencia de personas")

# ================= DIA 16: CONCLUSIÓN =================
slide16 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_background(slide16, FONDO_NEGRO)
add_centered_title(slide16, "Diseñemos un mundo")
add_centered_title(slide16, "con fila cero")
# Ajustar segundo título
tb3 = slide16.shapes.add_textbox(Inches(1), Inches(3.5), Inches(8), Inches(1))
p3 = tb3.text_frame.paragraphs[0]
p3.text = "con fila cero"
p3.font.size = Pt(44)
p3.font.bold = True
p3.font.color.rgb = DORADO
p3.alignment = PP_ALIGN.CENTER
cta = slide16.shapes.add_textbox(Inches(2), Inches(5), Inches(6), Inches(1))
cta.text_frame.text = "Invierte en TurnoUni: eficiencia operativa + transformación digital"
cta.text_frame.paragraphs[0].font.size = Pt(20)
cta.text_frame.paragraphs[0].font.color.rgb = GRIS_CLARO
cta.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

# Guardar
prs.save("TurnoUni_Final.pptx")
print("✅ Presentación generada: TurnoUni_Final.pptx")
