from shiny import App, ui, render, reactive
import pandas as pd
import tempfile
import os
from pathlib import Path

# =========================
# UI
# =========================
app_ui = ui.page_fluid(

    # =========================
    # ESTILOS + FUENTES
    # =========================
    ui.tags.head(
        ui.tags.link(
            href="https://fonts.googleapis.com/css2?family=DM+Sans&family=Space+Grotesk:wght@500;700&display=swap",
            rel="stylesheet"
        ),
        ui.tags.style("""
        :root {
            --font-main: 'DM Sans', sans-serif;
            --font-title: 'Space Grotesk', sans-serif;

            --bg: #fff9f5;
            --text: #2d3748;
            --primary: #22c55e;
            --secondary: #e0f2fe;
            --header: #53a689;

            --shadow: 0 4px 10px rgba(0,0,0,0.08);
        }

        body {
            background-color: var(--bg);
            font-family: var(--font-main);
            color: var(--text);
        }

        h1, h2, h3 {
            font-family: var(--font-title);
            color: var(--header);
        }

        .header-box {
            text-align: center;
            padding: 20px;
            background: white;
            border-radius: 12px;
            box-shadow: var(--shadow);
            margin-bottom: 20px;
        }

        .card {
            background: white;
            padding: 15px;
            border-radius: 12px;
            box-shadow: var(--shadow);
            margin-bottom: 15px;
        }

        table {
            margin: auto;
            text-align: center;
        }

        th, td {
            text-align: center !important;
        }

        .footer {
            text-align: center;
            font-size: 12px;
            color: gray;
            margin-top: 30px;
        }

        .footer img {
            height: 24px;
            margin: 0 10px;
        }

        .note-box {
            background: #f0fdf4;
            border-left: 5px solid var(--primary);
        }
        """)
    ),

    # =========================
    # HEADER
    # =========================
    ui.div(
        ui.img(src="logo.png", height="90px"),
        ui.h2("Sistema de Cálculo de Volumen Forestal"),
        ui.p("Procesamiento de censos y generación de matrices diamétricas"),
        class_="header-box"
    ),

    # =========================
    # CONTROLES
    # =========================
    ui.div(
        ui.input_file("file", "Sube Excel de campo (.xlsx)"),
        ui.input_action_button("run", "Procesar"),
        ui.input_action_button("reset", "🔄 Reiniciar"),
        ui.download_button("descargar", "⬇️ Descargar Excel"),
        class_="card"
    ),

    # =========================
    # PREVIEW
    # =========================
    ui.div(
        ui.h4("Vista previa"),
        ui.output_table("preview"),
        class_="card"
    ),

    # =========================
    # NOTA TÉCNICA
    # =========================
    ui.div(
        ui.h5("Nota técnica"),
        ui.p("Ecuación genérica para el cálculo de volumen en el género Pinus."),
        ui.tags.code("VT = (0.0115833000 + 0.0000440250 × (Dn²) × AT)"),
        ui.p("Donde Dn = diámetro (cm) y AT = altura (m)."),
        class_="card note-box"
    ),

    # =========================
    # ESTADO
    # =========================
    ui.div(
        ui.output_text_verbatim("estado_txt"),
        class_="card"
    ),

    # =========================
    # FOOTER
    # =========================
    ui.div(
        ui.p("Desarrollado por José Ricardo Eguiz Portillo | 2026 ©"),
        ui.a(ui.img(src="gmail.png"), href="mailto:tu_correo@gmail.com"),
        ui.a(ui.img(src="outlook.png"), href="mailto:tu_correo@outlook.com"),
        class_="footer"
    )
)

# =========================
# SERVER
# =========================
def server(input, output, session):

    estado = reactive.Value("⬆️ Sube un archivo para comenzar")
    preview_data = reactive.Value(None)
    ruta_archivo = {"ruta": None}

    @reactive.Effect
    @reactive.event(input.reset)
    def _():
        preview_data.set(None)
        ruta_archivo["ruta"] = None
        estado.set("🔄 Sistema reiniciado")

    @reactive.Effect
    @reactive.event(input.run)
    def _():

        if not input.file():
            estado.set("⚠️ Sube un archivo Excel")
            return

        try:
            df = pd.read_excel(input.file()[0]["datapath"])

            df.columns = df.columns.str.strip().str.lower().str.replace("#", "")

            cols = ["coordenada","especie","diametro","altura","agente causal"]
            if not all(c in df.columns for c in cols):
                estado.set("❌ Columnas incorrectas")
                return

            df["categoria diametrica"] = round(df["diametro"]/5)*5
            df["categoria altura"] = round(df["altura"]/5)*5

            df["volindividual"] = 0.0115833 + 0.000044025*(df["diametro"]**2)*df["altura"]
            df["volindividualcat"] = 0.0115833 + 0.000044025*(df["categoria diametrica"]**2)*df["categoria altura"]

            df["voltotal"] = df["volindividual"].sum()
            df["volcat"] = df["volindividualcat"].sum()

            matriz_conteo = pd.pivot_table(df, index="categoria diametrica",
                                           columns="categoria altura",
                                           values="diametro", aggfunc="count", fill_value=0)

            matriz_volumen = pd.pivot_table(df, index="categoria diametrica",
                                            columns="categoria altura",
                                            values="volindividualcat", aggfunc="sum", fill_value=0)

            tmp = tempfile.mkdtemp()
            ruta = os.path.join(tmp, "resultado.xlsx")

            with pd.ExcelWriter(ruta) as writer:
                df.to_excel(writer, sheet_name="Datos", index=False)
                matriz_conteo.to_excel(writer, sheet_name="Conteo")
                matriz_volumen.to_excel(writer, sheet_name="Volumen")

            ruta_archivo["ruta"] = ruta

            preview_data.set(df.head(10))
            estado.set("✅ Proceso completado")

        except Exception as e:
            estado.set(f"❌ Error: {str(e)}")

    @output
    @render.table
    def preview():
        return preview_data.get() if preview_data.get() is not None else pd.DataFrame()

    @output
    @render.text
    def estado_txt():
        return estado.get()

    @output
    @render.download(filename="resultado.xlsx")
    def descargar():
        return ruta_archivo["ruta"]

# =========================
app = App(app_ui, server, static_assets=Path(__file__).parent / "images")