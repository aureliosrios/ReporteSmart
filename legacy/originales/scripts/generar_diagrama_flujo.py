import matplotlib.pyplot as plt
import matplotlib.patches as patches

def draw_flowchart():
    fig, ax = plt.subplots(figsize=(12, 10), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    fig.patch.set_facecolor('#0f172a') # Fondo oscuro elegante
    ax.set_facecolor('#0f172a')

    # Estilos de Cajas (Glassmorphism & Colors)
    color_c1 = '#1e293b' # Captura (Dark Blue Box)
    color_c2 = '#1e1b4b' # Procesamiento (Dark Indigo)
    color_c3 = '#31103f' # Analítica / IA (Dark Purple)
    color_c4 = '#1c3d5a' # Gobernanza (Dark Cyan)

    border_c1 = '#38bdf8' # Accent Blue
    border_c2 = '#818cf8' # Accent Indigo
    border_c3 = '#c084fc' # Accent Purple
    border_c4 = '#4ade80' # Accent Green

    # Título Principal
    ax.text(50, 96, "DIAGRAMA DE FLUJO OPERATIVO: CONTROL DEL RESULTADO OPERATIVO (RO) CON IA",
            color='#ffffff', fontsize=12, fontweight='bold', ha='center', va='center', family='sans-serif')
    ax.text(50, 93, "Proyecto Redes Sanitarias (60 Días) | Captura Simplificada en 4 Nodos WBS",
            color='#94a3b8', fontsize=9, ha='center', va='center', family='sans-serif')

    # 1. BLOQUE CAPTURA DIARIA (Fila 1)
    rect1 = patches.FancyBboxPatch((5, 68), 90, 21, boxstyle="round,pad=1,rounding_size=2",
                                  facecolor=color_c1, edgecolor=border_c1, linewidth=2)
    ax.add_patch(rect1)
    ax.text(50, 86, "1. CAPTURA DIARIA EN OBRA (Cierre 6:00 PM)", color='#38bdf8', fontsize=10, fontweight='bold', ha='center')

    # Subcajas Inputs
    inputs = [
        ("Tareador (Bildin)", "Horas Hombre (HH)\npor WBS", 14),
        ("Almacenero (Vales)", "Materiales Despachados\npor WBS", 38),
        ("Administradora", "Partes Equipos (HM)\npor WBS", 62),
        ("Ing. de Campo", "Metrados Ejecutados\npor Partida", 86)
    ]
    for title, desc, x_pos in inputs:
        b = patches.FancyBboxPatch((x_pos-10, 71), 20, 11, boxstyle="round,pad=0.5,rounding_size=1",
                                   facecolor='#0f172a', edgecolor='#334155', linewidth=1.5)
        ax.add_patch(b)
        ax.text(x_pos, 79, title, color='#f8fafc', fontsize=8, fontweight='bold', ha='center')
        ax.text(x_pos, 74, desc, color='#94a3b8', fontsize=7, ha='center')

    # Flecha 1 -> 2
    ax.annotate('', xy=(50, 48), xytext=(50, 66),
                arrowprops=dict(arrowstyle="->", color='#38bdf8', lw=2.5))
    ax.text(52, 57, "Consolidación de Datos Acumulados", color='#38bdf8', fontsize=8, fontweight='bold')

    # 2. BLOQUE PROCESAMIENTO ETL (Fila 2)
    rect2 = patches.FancyBboxPatch((15, 42), 70, 15, boxstyle="round,pad=1,rounding_size=2",
                                  facecolor=color_c2, edgecolor=border_c2, linewidth=2)
    ax.add_patch(rect2)
    ax.text(50, 54, "2. MOTOR DE PROCESAMIENTO & EVM (Viernes 7:00 PM)", color='#818cf8', fontsize=10, fontweight='bold', ha='center')

    b_ac = patches.FancyBboxPatch((20, 44), 18, 7, boxstyle="round,pad=0.5,rounding_size=1", facecolor='#0f172a', edgecolor='#818cf8')
    ax.add_patch(b_ac)
    ax.text(29, 47.5, "AC (Costo Real)\n= MO + MAT + EQP", color='#f8fafc', fontsize=7.5, fontweight='bold', ha='center')

    b_ev = patches.FancyBboxPatch((41, 44), 18, 7, boxstyle="round,pad=0.5,rounding_size=1", facecolor='#0f172a', edgecolor='#4ade80')
    ax.add_patch(b_ev)
    ax.text(50, 47.5, "EV (Valor Ganado)\n= Metrado × PU Meta", color='#f8fafc', fontsize=7.5, fontweight='bold', ha='center')

    b_pv = patches.FancyBboxPatch((62, 44), 18, 7, boxstyle="round,pad=0.5,rounding_size=1", facecolor='#0f172a', edgecolor='#38bdf8')
    ax.add_patch(b_pv)
    ax.text(71, 47.5, "PV (Planificado)\n= Cronograma × PU Meta", color='#f8fafc', fontsize=7.5, fontweight='bold', ha='center')

    # Flecha 2 -> 3
    ax.annotate('', xy=(50, 27), xytext=(50, 40),
                arrowprops=dict(arrowstyle="->", color='#c084fc', lw=2.5))
    ax.text(52, 33.5, "Cálculo de CPI, SPI y Diagnóstico de Brechas", color='#c084fc', fontsize=8, fontweight='bold')

    # 3. BLOQUE ANALÍTICA E IA (Fila 3)
    rect3 = patches.FancyBboxPatch((15, 17), 70, 15, boxstyle="round,pad=1,rounding_size=2",
                                  facecolor=color_c3, edgecolor=border_c3, linewidth=2)
    ax.add_patch(rect3)
    ax.text(50, 29, "3. AGENTE IA DE BRECHAS & PROYECCIÓN EAC (Lunes 7:00 AM)", color='#c084fc', fontsize=10, fontweight='bold', ha='center')

    b_ia = patches.FancyBboxPatch((20, 19), 60, 7, boxstyle="round,pad=0.5,rounding_size=1", facecolor='#0f172a', edgecolor='#c084fc')
    ax.add_patch(b_ia)
    ax.text(50, 22.5, "Evaluación de CPI < 1.0 por WBS -> Proyección EAC -> Informe Sintético de Causa Raíz",
            color='#f8fafc', fontsize=8, ha='center')

    # Flecha 3 -> 4
    ax.annotate('', xy=(50, 8), xytext=(50, 15),
                arrowprops=dict(arrowstyle="->", color='#4ade80', lw=2.5))

    # 4. BLOQUE GOBERNANZA (Fila 4)
    rect4 = patches.FancyBboxPatch((20, 1), 60, 9, boxstyle="round,pad=1,rounding_size=2",
                                  facecolor=color_c4, edgecolor=border_c4, linewidth=2)
    ax.add_patch(rect4)
    ax.text(50, 6.5, "4. GOBIERNO DE OBRA (Reunión Semanal Lunes 8:30 AM)", color='#4ade80', fontsize=9.5, fontweight='bold', ha='center')
    ax.text(50, 3, "Residente de Obra + Ing. Campo + OT -> Decisiones y Mitigación de Sobrecostos", color='#cbd5e1', fontsize=8, ha='center')

    plt.tight_layout()
    img_path = "d:\\Agentes de IA\\Habilitación urbana\\diagrama_flujo_ro.png"
    plt.savefig(img_path, facecolor=fig.get_facecolor(), edgecolor='none', bbox_inches='tight')
    plt.close()
    print(f"Imagen del Diagrama de Flujo generada exitosamente en: {img_path}")

if __name__ == "__main__":
    draw_flowchart()
