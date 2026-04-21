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
    # ESTILOS
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
            --header: #53a689;
        }

        body {
            background-color: var(--bg);
            font-family: var(--font-main);
            color: var(--text);
        }

        h2, h3 { font-family: var(--font-title); color: var(--header); }

        .card {
            background: white;
            padding: 15px;
            border-radius: 12px;
            margin-bottom: 15px;
        }

        table, th, td {
            text-align: center !important;
            margin: auto;
        }

        .footer {
            text-align: center;
            font-size: 12px;
            color: gray;
        }

        .note-box {
            background: #f0fdf4;
            border-left: 5px solid var(--primary);
            padding: 10px;
        }
        """)
    ),

    # =========================
    # HEADER
    # =========================
    ui.div(
        ui.img(
            src="logo.png",
            style="max-height:70px; width:auto; display:block; margin:auto;"
        ),
        ui.h2("Sistema de Cálculo de Volumen Forestal"),
        ui.p("Procesamiento de censos de campo"),
        class_="card",
        style="text-align:center;"
    ),

    # =========================
    # CONTROLES
    # =========================
    ui.div(
        ui.input_file("file", "Sube Excel (.xlsx)"),
        ui.input_action_button("run", "Procesar"),
        ui.input_action_button("reset", "🔄 Reiniciar"),
        ui.download_button("descargar", "⬇️ Descargar"),
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
    # NOTA TÉCNICA + FORMATO
    # =========================
    ui.div(
        ui.h5("📌 Nota técnica y formato de entrada"),

        ui.p("Ecuación utilizada para el cálculo de volumen (género Pinus):"),
        ui.tags.code("VT = (0.0115833000 + 0.0000440250 × (Dn²) × AT)"),

        ui.p("Donde:"),
        ui.tags.ul(
            ui.tags.li("VT = Volumen total (m³ vta)"),
            ui.tags.li("Dn = Diámetro normal (cm)"),
            ui.tags.li("AT = Altura total (m)")
        ),

        ui.p("Esta es una ecuación genérica para Pinus."),

        ui.hr(),

        ui.h6("📋 Formato esperado del archivo Excel"),
        ui.tags.code("Coordenada | Especie | Diametro | Altura | Agente Causal"),

        ui.p("Ejemplo de captura:"),
        ui.tags.pre(
            "1 | Pinus arizonica | 24 | 13 | Dendroctonus mexicanus\n"
            "2 | Pinus arizonica | 10 | 5 | Dendroctonus mexicanus"
        ),

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
        ui.img(src="gmail.png", height="20px"),
        ui.img(src="outlook.png", height="20px"),
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

    # RESET
    @reactive.Effect
    @reactive.event(input.reset)
    def _():
        preview_data.set(None)
        ruta_archivo["ruta"] = None
        estado.set("🔄 Sistema reiniciado")

    # PROCESO
    @reactive.Effect
    @reactive.event(input.run)
    def _():

        if not input.file():
            estado.set("⚠️ Debes subir un archivo Excel")
            return

        try:
            df = pd.read_excel(input.file()[0]["datapath"])

            # NORMALIZAR
            def norm(x):
                return str(x).strip().lower()

            df.columns = [norm(c) for c in df.columns]

            alias = {
                "coordenada": ["coordenada","coord","punto"],
                "especie": ["especie","sp"],
                "diametro": ["diametro","dbh"],
                "altura": ["altura","h"],
                "agente causal": ["agente causal","plaga"]
            }

            mapeo = {}
            for k, vals in alias.items():
                for col in df.columns:
                    if col in vals:
                        mapeo[k] = col

            faltantes = [k for k in alias if k not in mapeo]

            if faltantes:
                estado.set(
                    "❌ Formato incorrecto del archivo\n\n"
                    "📌 Debe contener:\n"
                    "Coordenada | Especie | Diametro | Altura | Agente Causal\n\n"
                    f"⚠️ Faltan: {faltantes}\n"
                    f"🔍 Detectadas: {list(df.columns)}\n\n"
                    "💡 Revisa nombres o encabezados"
                )
                return

            df = df.rename(columns={v:k for k,v in mapeo.items()})

            # =========================
            # CÁLCULOS
            # =========================
            df["categoria diametrica"] = round(df["diametro"]/5)*5
            df["categoria altura"] = round(df["altura"]/5)*5

            df["volindividual"] = 0.0115833 + 0.000044025*(df["diametro"]**2)*df["altura"]
            df["volindividualcat"] = 0.0115833 + 0.000044025*(df["categoria diametrica"]**2)*df["categoria altura"]

            df["voltotal"] = df["volindividual"].sum()
            df["volcat"] = df["volindividualcat"].sum()

            # MATRIZ
            matriz = pd.pivot_table(
                df,
                index="categoria diametrica",
                columns="categoria altura",
                values="diametro",
                aggfunc="count",
                fill_value=0
            )

            # EXPORTAR
            tmp = tempfile.mkdtemp()
            ruta = os.path.join(tmp, "resultado.xlsx")

            with pd.ExcelWriter(ruta) as writer:
                df.to_excel(writer, index=False)
                matriz.to_excel(writer, sheet_name="Matriz")

            ruta_archivo["ruta"] = ruta

            preview_data.set(df.head(10))
            estado.set("✅ Procesado correctamente")

        except Exception as e:
            estado.set(f"❌ Error: {str(e)}")

    # OUTPUTS
    @output
    @render.table
    def preview():
        if preview_data.get() is None:
            return pd.DataFrame({"Info": ["Sin datos aún"]})
        return preview_data.get()

    @output
    @render.text
    def estado_txt():
        return estado.get()

    @output
    @render.download(filename="resultado.xlsx")
    def descargar():
        return ruta_archivo["ruta"]

# =========================
# APP
# =========================
app = App(app_ui, server, static_assets=Path(__file__).parent / "images")