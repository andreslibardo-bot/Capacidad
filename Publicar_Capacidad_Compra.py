import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import io
from datetime import datetime
import sys

import Capacidad_compra as cc

st.write(sys.executable)

st.title("Modelo de Capacidad de compra")

st.markdown("### ⚙️ Tipo de proceso")

tipo_proceso = st.radio( "Seleccione el proceso a ejecutar:",
        [
            "Cilindros y Tanques",
            "Cilindros, Tanques y Redes"
        ]
    )

st.markdown("### 📂 Cargar archivos")

file_Cilindros = st.file_uploader( "📄 Archivo y hoja de cilindros", type=["xlsx"])

file_tanques = st.file_uploader( "📄 Archivo y hojas de tanques", type=["xlsx"])

file_Capac_anterior = st.file_uploader( "📄 Archivo capacidad periodo anterior (opcional)",  type=["xlsx"])

# Solo si el usuario selecciona el proceso completo
file_Redes_SUI = None
file_Redes_Proyec = None

if tipo_proceso == "Cilindros, Tanques y Redes":

    file_Redes_SUI = st.file_uploader( "📄 Archivo y hoja de ventas en Redes SUI",  type=["xlsx"] )

    file_Redes_Proyec = st.file_uploader( "📄 Archivo y hoja de Proyección mercado de redes", type=["xlsx"] )


if tipo_proceso == "Cilindros y Tanques":

    archivos_ok = (
        file_Cilindros is not None and
        file_tanques is not None
    )

else:

    archivos_ok = (
        file_Cilindros is not None and
        file_tanques is not None and
        file_Redes_SUI is not None and
        file_Redes_Proyec is not None
    )

if archivos_ok:
    st.success("Archivos cargados correctamente ✔")

    # -------------------------
    # Cilindros
    # -------------------------

    hojas_cilindros = cc.obtener_hojas_excel(file_Cilindros)
    hoj_cil = st.selectbox( "Selecciona la hoja de cilindros", hojas_cilindros )
    df_cilindros = cc.leer_hoja_excel( file_Cilindros, hoj_cil )

    # -------------------------
    # Tanques
    # -------------------------

    hojas_tanques = cc.obtener_hojas_excel(file_tanques)
    hoja_tanques_T_periodo = st.selectbox( "Selecciona hoja Tanques Tm", hojas_tanques)
    hoja_tanques_T_menos_1_periodo = st.selectbox( "Selecciona hoja Tanques Tm-1", hojas_tanques)
    df_tanques_TM = cc.leer_hoja_excel(file_tanques, hoja_tanques_T_periodo)
    df_tanques_TM_1 = cc.leer_hoja_excel(file_tanques, hoja_tanques_T_menos_1_periodo)

    # -------------------------
    # Redes y Proyecciones
    # -------------------------

    if tipo_proceso == "Cilindros, Tanques y Redes":

        # Redes
        hojas_redes = cc.obtener_hojas_excel(file_Redes_SUI)

        hoj_red = st.selectbox( "Selecciona la hoja de redes",  hojas_redes  )

        df_red = cc.leer_hoja_excel( file_Redes_SUI, hoj_red  )

        # Proyecciones
        hojas_apligas = cc.obtener_hojas_excel(file_Redes_Proyec)

        hoja_proyec_usuar_demand = st.selectbox( "Selecciona hoja proyec_usuar",  hojas_apligas,  index=hojas_apligas.index("Usuarios_Demanda_nueva")  )

        hoja_inversio_nuev = st.selectbox( "Selecciona hoja Inversion_nueva", hojas_apligas, index=hojas_apligas.index("Inversion_nueva")  )

        df_proyec_usuar_demand = cc.leer_hoja_excel( file_Redes_Proyec, hoja_proyec_usuar_demand )

        df_inversio_nuev = cc.leer_hoja_excel( file_Redes_Proyec, hoja_inversio_nuev )

        # Mes de inicio
        meses = { "Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4,  "Mayo": 5, "Junio": 6,  "Julio": 7, "Agosto": 8, "Septiembre": 9, "Octubre": 10,  "Noviembre": 11, "Diciembre": 12  }

        mes_seleccionado = st.selectbox(
            "Seleccione el mes de inicio",
            list(meses.keys())
        )

        inicio_periodo = meses[mes_seleccionado]

    # Mostrar preview
    with st.expander("👀 Ver datos cargados"):

        st.write("Cilindros", df_cilindros.head())
        st.write("Tanques_TM", df_tanques_TM.head())
        st.write("Tanques_TM-1", df_tanques_TM_1.head())

        if tipo_proceso == "Cilindros, Tanques y Redes":

            st.write("Red", df_red.head())
            st.write( "Proyección usuarios demand",  df_proyec_usuar_demand.head()   )
            st.write( "Inversión nueva", df_inversio_nuev.head()  )
            st.write( "Mes seleccionado:",  inicio_periodo  )
    

    # -------------------------
    # Botón Calcular
    # -------------------------
    if "model_results" not in st.session_state:
        st.session_state.model_results = None

    if st.button("🚀 Ejecutar Capacidad de compra"):
        # Spinner (UX)
        with st.spinner("Ejecutando modelo ..."):

            #######################################################   ejecutar  funciones ##############################

            if tipo_proceso == "Cilindros y Tanques":
                st.write("Ejecutando cilindros y tanques")

                resultados = cc.ejecutar_cilindros_tanques(
                    df_cilindros,
                    df_tanques_TM,
                    df_tanques_TM_1
                )

            else:
                    
                resultados = cc.ejecutar_completo(
                    df_cilindros,
                    df_tanques_TM,
                    df_tanques_TM_1,
                    df_inversio_nuev,
                    df_proyec_usuar_demand,
                    df_red,
                    inicio_periodo
                )
   
                reportes = cc.generar_reportes(
                    resultados["Capacidad_GLP"],
                    resultados["resultados_cil"],
                    resultados["resultados_tan"],
                    resultados["Cap_Mer_Ini_consol"],
                    resultados["Cap_Mer_Ope"]
                )

                resultados.update(reportes)

            resultados["file_Capac_anterior"] = file_Capac_anterior

            st.session_state.model_results = resultados   
        
            st.success("Modelo ejecutado correctamente")

    # ==========================
    # Recuperar resultados
    # ==========================
    if st.session_state.model_results is not None:

        resultados = st.session_state.model_results

        st.write(
        f"Cálculos realizados usando la información proveniente del SUI de "
        f"{resultados['num_empresas']} distribuidores."
        )

        st.text(f"""Información cilindros
        Empresas operativas con cilindros: {resultados['empresas_cilindros']}
        Cantidad de marcas con registros: {resultados['marcas_cilindros']}
        Cantidad de cilindros: {resultados['cilindros_total']:,.2f}
        """)

        st.text(f"""Información de tanques estacionarios
        Empresas operativas con tanques Trimestre m-1: {resultados['empresas_tanques_T3']}
        Cantidad de tanques Trimestre m-1: {resultados['tanques_T3']:,.2f}
        Empresas operativas con tanques Trimestre m: {resultados['empresas_tanques_T4']}
        Cantidad de tanques Trimestre m: {resultados['tanques_T4']:,.2f}
        """)
    
        # -------------------------
        # Descargar Word
        # -------------------------
        cap_glp = resultados["Capacidad_GLP_tot"]

        if "Código_SUI_num" in cap_glp.columns:
            cap_glp = cap_glp.drop(columns=["Código_SUI_num"])

        tablas = [
                (
                    "capacidad de envase en cilindros, para cada marca de propiedad del distribuidor de acuerdo con la información registrada desde el 2008 hasta octubre de 2012, por AIC proyectos",
                    resultados["Cilindros_LB"].drop(columns=["Empresa"])
                ),
                (
                    "Capacidad de envase en cilindros, para cada marca de propiedad del distribuidor de acuerdo con la información registrada al SUI desde noviembre de 2012 hasta la fecha",
                    resultados["Cilindros_KG"].drop(columns=["Empresa"])
                ),
                (
                    "Capacidad total de envase en cilindros de propiedad del distribuidor",
                    resultados["Cilindros_consol"]
                ),
                (
                    "La capacidad total de tanques estacionarios atendidos por el distribuidor",
                    resultados["resultados_tanques"]
                ),
                (
                    "La capacidad de tanques y cilindros estacionarios atendidos por el distribuidor",
                    resultados["resultado_CC_tanques"]
                ),
                (
                    "La capacidad de compra de cada distribuidor",
                    cap_glp
                ),
            ]
        
       
        if all(k in resultados for k in [
            "Cap_Mer_Ini_consol",
            "Cap_Mer_Opera_cons",
            "Capacidad_Redes"
        ]):

            tablas.extend([
                (
                    "La capacidad de mercado inicial de cada distribuidor",
                    resultados["Cap_Mer_Ini_consol"]
                ),
                (
                    "La capacidad de mercado en operación de cada distribuidor",
                    resultados["Cap_Mer_Opera_cons"]
                ),
                (
                    "La capacidad de redes de cada distribuidor",
                    resultados["Capacidad_Redes"]
                ),
                (
                    "La capacidad de compra de cada distribuidor",
                    cap_glp
                )
            ])

 
        word_file = cc.generar_word_completo(tablas)

        fecha_hora = datetime.now().strftime("%Y-%m-%d_%H-%M")  # formato: 2026-05-22_14-30
        
        nombre_archivo_word = f"Reporte_Tablas_word {fecha_hora}.docx" 

        st.download_button(
            label="Descargar en Word",
            data=word_file,
            file_name=nombre_archivo_word,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        # -------------------------
        # Descargar Excel
        # -------------------------
       
        buffer = io.BytesIO()

        hojas_excel = {
            "Cilindros_LB": "Cilindros_LB",
            "Cilindros_KG": "Cilindros_KG",
            "Cilindros_consol": "Cilindros_consol",
            "resultados_tanques": "resultados_tanques",
            "resultado_CC_tanques": "resultados_cilindros_tanques",
            "validacion_tipo_red": "validacion_tipo_red",
            "mercados_aplig_falta": "mercados_aplig_falta",
            "Cap_Mer_Ini_base": "Cap_Mer_Ini_base",
            "Cap_Mer_Ini_consol": "Cap_Mer_Ini_consol",
            "Cap_Mer_Opera_cons": "Cap_Mer_Ope_cons",
            "Cap_Mer_Ope": "Cap_Mer_Ope",     
            "Capacidad_Redes": "Capacidad_Redes",
            "Capacidad_GLP_tot": "Capacidad_GLP_tot",
            "participa_capacidad": "participa_capacidad"
        }

        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:

            for clave, hoja in hojas_excel.items():

                if clave in resultados:

                    resultados[clave].to_excel(
                        writer,
                        sheet_name=hoja,
                        index=False
                    )

        nombre_archivo_excel = f"Tablas_intermedias_CCit_{fecha_hora}.xlsx"

        st.download_button(
            label="📥 Descargar en Excel",
            data=buffer.getvalue(),
            file_name=nombre_archivo_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # -------------------------
        # Comparacion con modeloa anterior
        # -------------------------
       
        if resultados["file_Capac_anterior"] is not None:
            actual = cc.preparar_comparacion(resultados)

            anterior_raw = pd.read_excel( resultados["file_Capac_anterior"], sheet_name=None  )

            anterior = {
                hoja: cc.estandarizar_llave(df)
                for hoja, df in anterior_raw.items()
            }


            comparaciones = cc.comparar_workbooks(  actual,  anterior,  key="id_empresa",  columnas_ignorar=[
                                     "Código SUI / Capacidad",    "Empresa",    "Código SUI", "id_empresa",
                                      "EMPRESA",  "ID_SUI",  "RAZON_SOCIAL",  "ID_EMPRESA", "Prestador",
                                      "valid_A",  "valid_B", "valida_total", "Código_SUI_num", "EMPRES", "indice_empresa"
                                       ]
                                       )

            # para exportar la comparacion con el periodo anterior
            buffer_comp = io.BytesIO()

            with pd.ExcelWriter(buffer_comp, engine="openpyxl") as writer:

                for hoja, df in comparaciones.items():
                    df.to_excel(writer, sheet_name=f"comp_{hoja}"[:31], index=False)

            buffer_comp.seek(0)

            st.download_button(
            "📊 Descargar comparación",
            data=buffer_comp.getvalue(),
            file_name=f"comparacion_{fecha_hora}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # ==========================
            # Alertas variaciones > 20%
            # ==========================

            alertas = []

            for hoja, df_comp in comparaciones.items():

                # columnas de porcentaje
                cols_pct = [c for c in df_comp.columns if c.startswith("pct_")]

                for col in cols_pct:

                    # filas con variación mayor al 20%
                    df_alerta = df_comp[
                        df_comp[col].abs() > 0.20
                    ].copy()

                    if not df_alerta.empty:

                        nombre_variable = col.replace("pct_", "")

                        for _, fila in df_alerta.iterrows():

                            empresa = fila.get("id_empresa", "Sin ID")
                            variacion = fila[col]

                            tipo = "Incremento" if variacion > 0 else "Disminución" 

                            alertas.append({
                                "Hoja": hoja,
                                "Empresa": empresa,
                                "Variable": nombre_variable,
                                "tipo": tipo,
                                "Variacion_%": variacion
                            })

            # ==========================
            # Mostrar alertas
            # ==========================

            if alertas:

                df_alertas = pd.DataFrame(alertas)

                df_alertas["Variacion_%"] = (
                    df_alertas["Variacion_%"] * 100
                ).round(2)

                st.warning(
                    f"⚠️ Se encontraron {len(df_alertas)} variaciones mayores al 20%"
                )

                st.dataframe(df_alertas)

            else:

                st.success("✅ No se encontraron variaciones mayores al 20%")
