import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import io
from datetime import datetime
import sys

import Capacidad_compra as cc

st.title("Modelo de Capacidad de compra")

st.markdown("### 📂 Cargar archivos")

# 🔽 Uploaders
file_Cilindros = st.file_uploader("📄 Archivo y hoja de cilindros", type=["xlsx"])
file_tanques = st.file_uploader("📄 Archivo y hojas de tanques", type=["xlsx"])
file_Redes_SUI = st.file_uploader("📄 Archivo y hoja de ventas en Redes SUI", type=["xlsx"])
file_Redes_Proyec = st.file_uploader("📄 Archivo y hoja de Proyeccion mercado de redes", type=["xlsx"])
file_Capac_anterior = st.file_uploader("📄 Archivo y hoja de Capacidad periodo Anterior", type=["xlsx"])

if file_Cilindros and file_tanques and file_Redes_SUI and file_Redes_Proyec:

    st.success("Archivos cargados correctamente ✔")

    ############################################ Leer archivos cilindro  ######
    
    hojas_cilindros = pd.ExcelFile(file_Cilindros).sheet_names
    
    # Permitir que el usuario seleccione la hoja
    hoj_cil = st.selectbox("Selecciona la hoja de cilindros", hojas_cilindros)
    
    # Leer la hoja seleccionada
    df_cilindros = pd.read_excel(file_Cilindros, sheet_name=hoj_cil)
    
    ############################################ Leer archivos tanques  ######

    # Obtener todas las hojas disponibles
    hojas_tanques = pd.ExcelFile(file_tanques).sheet_names

    # Selección de hojas por el usuario
    hoja_tanques_T_periodo = st.selectbox("Selecciona hoja Tanques Tm", hojas_tanques)
    hoja_tanques_T_menos_1_periodo = st.selectbox("Selecciona hoja Tanques Tm-1", hojas_tanques)

    # Leer las hojas seleccionadas
    df_tanques_TM = pd.read_excel(file_tanques, sheet_name=hoja_tanques_T_periodo)
    df_tanques_TM_1 = pd.read_excel(file_tanques, sheet_name=hoja_tanques_T_menos_1_periodo)

    ########################################### Leer archivos redes  ######
    
    hojas_redes = pd.ExcelFile(file_Redes_SUI).sheet_names
    
    # Permitir que el usuario seleccione la hoja
    hoj_red = st.selectbox("Selecciona la hoja de redes", hojas_redes)
    
    # Leer la hoja seleccionada
    df_red = pd.read_excel(file_Redes_SUI, sheet_name=hoj_red)

    ########################################### Leer archivos proyecciones Apligas  ######

    # Obtener todas las hojas disponibles
    hojas_Apligas = pd.ExcelFile(file_Redes_Proyec).sheet_names

    # Selección de hojas por el usuario
    hoja_proyec_usuar_demand = st.selectbox("Selecciona hoja proyec_usuar", hojas_Apligas, index=hojas_Apligas.index("Usuarios_Demanda_nueva"))
    hoja_inversio_nuev = st.selectbox("Selecciona hoja Inversion_nueva", hojas_Apligas, index=hojas_Apligas.index("Inversion_nueva"))

    # Leer las hojas seleccionadas
    df_proyec_usuar_demand = pd.read_excel(file_Redes_Proyec, sheet_name=hoja_proyec_usuar_demand)
    df_inversio_nuev = pd.read_excel(file_Redes_Proyec, sheet_name=hoja_inversio_nuev)
    
    meses = { "Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4,  "Mayo": 5, "Junio": 6,  "Julio": 7, "Agosto": 8, "Septiembre": 9, "Octubre": 10,  "Noviembre": 11, "Diciembre": 12 }

    mes_seleccionado = st.selectbox( "Seleccione el mes de inicio",  list(meses.keys())  )

    inicio_periodo = meses[mes_seleccionado]

    # Mostrar preview
    with st.expander("👀 Ver datos cargados"):
        st.write("Cilindros", df_cilindros.head())
        st.write("Tanques_TM", df_tanques_TM.head())
        st.write("Tanques_TM-1", df_tanques_TM_1.head())
        st.write("Red", df_red.head())
        st.write("Proyección usuarios demand", df_proyec_usuar_demand.head())
        st.write("Inversión nueva", df_inversio_nuev.head())
        st.write("Mes seleccionado:", inicio_periodo)

    
    # -------------------------
    # Botón Calcular
    # -------------------------
    if "model_results" not in st.session_state:
        st.session_state.model_results = None

    if st.button("🚀 Ejecutar Capacidad de compra"):

        # Spinner (UX)
        with st.spinner("Ejecutando modelo completo..."):
            #######################################################   ejecutar  funciones ##############################
            # 1. funcion General
            General = cc.calcular_CCit( df_cilindros, df_tanques_TM, df_tanques_TM_1)
                     
            # 2. procesamiento de los cilindros
            resultados_cil = cc.procesar_cilindros(df_cilindros, True)
          
            # 3. procesamiento de los tanques 
            resultados_tan = cc.procesar_tanques( df_tanques_TM, df_tanques_TM_1, True )

            # 4. procesamiento de la capacidad de Mercado inicial 
            Cap_Mer_Ini = cc.procesar_activos_completo(df_inversio_nuev, df_proyec_usuar_demand, True)
           
            # 5. procesamiento de la capacidad de Mercado en  operacion
           
            Cap_Mer_Ope = cc.calcular_indice_mensual_continuo(df_red, inicio_periodo)

            Cap_Mer_Opera_cons = (Cap_Mer_Ope.groupby(["ID_EMPRESA", "Prestador"], as_index=False)["Cap_Mer_Opera"].sum())

            # 6. procesamiento de la Capacidad Redes 
            
            Capacidad_Redes = cc.consolidar_activos(Cap_Mer_Ini["consolidado"], Cap_Mer_Ope)

            # 7. procesamiento de la Capacidad total cilindros, tanque y redes
            Capacidad_GLP = cc.consolidar_capacidades(resultados_tan['Final'],  resultados_cil['data_Cilindros_CapCil'],  Cap_Mer_Ini["consolidado"],  Cap_Mer_Ope, True) 
                       
            ####################################################### tablas Resolución ##############################
            #Capitulo 1 Cilindros LB
            Cilindros_LB = resultados_cil['cilindros_lb']

            Cilindros_LB  = Cilindros_LB.rename(columns={
                            "Código SUI": "Código SUI / Capacidad",
                            "Total_Cilindros": "Total Cilindros",
                            "Capacidad_LB": "Capacidad (Libras)"
                        })

            #Capitulo 2 Cilindros Kg
            Cilindros_KG = resultados_cil['Cilindros_KG']

            Cilindros_KG = Cilindros_KG.rename(columns={
                            "Código SUI": "Código SUI / Capacidad",
                            "Total_Cilindros": "Total general",
                            "Capacidad_kg": "Capacidad (kg)"
                        })

            # Cap_3_Cilindros_consolidado Kg
            Cilindros_consol = resultados_cil['data_Cilindros_CapCil']

            Cilindros_consol = Cilindros_consol.rename(columns={
                            "Cap_cil_kg": "Cap. cil (kg) "
                        })
            
            #Cap_4_Tanques
            resultados_tanques = resultados_tan['Final']

            resultados_tanques = resultados_tanques.rename(columns={
                            "Capacidad_gal": "Capacidad (gal)",
                            "Número_de_Tanques": "Número de Tanques",
                            "CapTEi_t_kg": "CapTEi,t (kg)"
                        })

            #Cap_5_tabla_final tanques y cilindros
            resultado_CC_tanques = General['resultado']

            resultado_CC_tanques = resultado_CC_tanques.rename(columns={
                            "CapTE": "CapTEi,t (kg)",
                            "Cap_cil": "Cap.cil (kg)",
                            "CCit": "(CapTE * 0.85 + Cap_cil)* 0.345)" })
            #Cap_6_Cap_Mer_Ini (indicada anteriormente)

            Cap_Mer_Ini_consol = Cap_Mer_Ini["consolidado"]

            #Cap_7_Cap_Mer_Ope (indicada anteriormente)

            #Cap_8_Capacidad_Redes (indicada anteriormente)

            #Cap_9_tabla_final tanques y cilindros y redes
            Capacidad_GLP_tot = Capacidad_GLP["merged_data_final"]
         
                  
            ####################################################### tablas presentacion ##############################
            # 1. funcion Empresas General
            empresas= cc.consolidar_agentes(Capacidad_GLP["merged_data_final"],   resultados_cil['Cilindros'],  resultados_tan,
                                         Cap_Mer_Ini_consol,  Cap_Mer_Ope, cc.limpiar_id, cc.limpiar_empresa,  True)

            # 2. funcion Resumen General
            Resum_gener = cc.visualizar_y_tabla_final(empresas["agentes_data"], Capacidad_GLP_tot, Capacidad_GLP["Capacidad_Redes"] ) 
            
            # 3. funcion Resumen General cilindros
            resumen = cc.resumen_tanques_cilindros(resultados_tan['T1_clean'], resultados_tan['T2_clean'], resultados_cil['Cilindros'])

            # 4. funcion Resumen participacion
            participa_capacidad = cc.crear_indice_empresas(empresas["agentes_data"], Capacidad_GLP_tot, cc.normalizar_nombre_empresa) 
            
            num_empresas = Resum_gener["Numero_empresas"]
            empresas_cilindros = resumen["empresas_cilindros"]
            marcas_cilindros = resumen["marcas_cilindros"]
            cilindros_total= resumen["cilindros_total"]
            empresas_tanques_T3= resumen["empresas_tanques_T3"]
            tanques_T3 = resumen["tanques_T3"]
            empresas_tanques_T4= resumen["empresas_tanques_T4"]
            tanques_T4= resumen["tanques_T4"]

            # Guardar resultados en session_state
            st.session_state.model_results = {
                "Cilindros_LB": Cilindros_LB,
                "Cilindros_KG": Cilindros_KG,
                "Cilindros_consol": Cilindros_consol,
                "resultados_tanques": resultados_tanques,
                "resultado_CC_tanques": resultado_CC_tanques,
                "Cap_Mer_Ini_consol": Cap_Mer_Ini_consol,
                "Cap_Mer_Opera_cons": Cap_Mer_Opera_cons, 
                "Capacidad_Redes": Capacidad_Redes,
                "Capacidad_GLP_tot": Capacidad_GLP_tot,
                "empresas":  empresas,
                "Resum_gener":  Resum_gener,
                "resumen":  resumen,
                "participa_capacidad":  participa_capacidad,    
                "num_empresas":  num_empresas ,
                "empresas_cilindros" : empresas_cilindros,
                "marcas_cilindros" : marcas_cilindros ,
                "cilindros_total"  : cilindros_total ,
                "empresas_tanques_T3"  : empresas_tanques_T3  ,
                "tanques_T3" : tanques_T3,
                "empresas_tanques_T4" : empresas_tanques_T4,
                "tanques_T4" : tanques_T4,
                "file_Capac_anterior" : file_Capac_anterior 
            }

    # ==========================
    # Recuperar resultados
    # ==========================
    if st.session_state.model_results is not None:

        Cilindros_LB = st.session_state.model_results["Cilindros_LB"]
        Cilindros_KG = st.session_state.model_results["Cilindros_KG"]
        Cilindros_consol = st.session_state.model_results["Cilindros_consol"]
        resultados_tanques = st.session_state.model_results["resultados_tanques"]
        resultado_CC_tanques = st.session_state.model_results["resultado_CC_tanques"]
        Cap_Mer_Ini_consol = st.session_state.model_results["Cap_Mer_Ini_consol"]
        Cap_Mer_Opera_cons = st.session_state.model_results["Cap_Mer_Opera_cons"]
        Capacidad_Redes = st.session_state.model_results["Capacidad_Redes"]
        Capacidad_GLP_tot = st.session_state.model_results["Capacidad_GLP_tot"]
        empresas = st.session_state.model_results["empresas"]
        Resum_gener = st.session_state.model_results["Resum_gener"]
        resumen = st.session_state.model_results["resumen"]
        participa_capacidad = st.session_state.model_results["participa_capacidad"]
        num_empresas = st.session_state.model_results["num_empresas"]
        empresas_cilindros = st.session_state.model_results["empresas_cilindros"]
        marcas_cilindros = st.session_state.model_results["marcas_cilindros"]
        cilindros_total = st.session_state.model_results["cilindros_total"]
        empresas_tanques_T3 = st.session_state.model_results["empresas_tanques_T3"]
        tanques_T3 = st.session_state.model_results["tanques_T3"]
        empresas_tanques_T4 = st.session_state.model_results["empresas_tanques_T4"]
        tanques_T4 = st.session_state.model_results["tanques_T4"]
        file_Capac_anterior = st.session_state.model_results["file_Capac_anterior"]
               
        st.write(f"Cálculos realizados usando la información proveniente del SUI de {num_empresas} distribuidores.")

        st.text(f""" Información cilindros
        Empresas operativas con cilindros:  {empresas_cilindros}
        Cantidad de marcas con registros: {marcas_cilindros}
        Cantidad de cilindros: {cilindros_total:,.2f}
        """)

        st.text(f""" Información de tanques estacionarios
        Empresas operativas con tanques Trimestre m-1:  {empresas_tanques_T3}
        Cantidad de tanques Trimetre m-1: {tanques_T3:,.2f}
        Empresas operativas con tanques Trimestre m: {empresas_tanques_T4}
        Cantidad de tanques Trimetre m: {tanques_T4:,.2f}
        """)
      
        # -------------------------
        # Descargar Word
        # -------------------------
        
        tablas = [
            ("capacidad de envase en cilindros, para cada marca de propiedad del distribuidor de acuerdo con la información registrada desde el 2008 hasta octubre de 2012, por AIC proyectos",  Cilindros_LB.drop(columns=["Empresa"])),
            ("Capacidad de envase en cilindros, para cada marca de propiedad del distribuidor de de acuerdo con la información registrada al SUI desde noviembre de 2012 hasta la fecha ", Cilindros_KG.drop(columns=["Empresa"])),
            ("Capacidad total de envase en cilindros de propiedad del distribuidor ", Cilindros_consol),
            ("La capacidad total de tanques estacionarios atendidos por el distribuidor ",  resultados_tanques),
            ("La capacidad de tanques y cilindros estacionarios atendidos por el distribuidor ",  resultado_CC_tanques),
            ("La capacidad de compra de cada distribuidor ", Capacidad_GLP_tot.drop(columns=["Código_SUI_num"])),
       ]
       
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

        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:

            Cilindros_LB.to_excel(writer, sheet_name="Cilindros_LB", index=False)
            Cilindros_KG.to_excel(writer, sheet_name="Cilindros_KG", index=False)
            Cilindros_consol.to_excel(writer, sheet_name="Cilindros_consol", index=False)
            resultados_tanques.to_excel(writer, sheet_name="resultados_tanques", index=False)
            resultado_CC_tanques.to_excel(writer, sheet_name="resultados_cilindros_tanques", index=False)
            Cap_Mer_Ini_consol.to_excel(writer, sheet_name="Cap_Mer_Ini", index=False)
            Cap_Mer_Opera_cons.to_excel(writer, sheet_name="Cap_Mer_Ope", index=False)
            Capacidad_Redes.to_excel(writer, sheet_name="Capacidad_Redes", index=False)
            Capacidad_GLP_tot.to_excel(writer, sheet_name="Capacidad_GLP_tot", index=False)
            participa_capacidad.to_excel(writer, sheet_name="participa_capacidad", index=False)

        nombre_archivo_excel = f"Tablas_intermedias_CCit_{fecha_hora}.xlsx"

        st.download_button(
            label="📥 Descargar en Excel",
            data=buffer.getvalue(),
            file_name = nombre_archivo_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
              
        if file_Capac_anterior is not None:
            anterior = pd.read_excel( file_Capac_anterior,  sheet_name=None  )

        actual = {
                "Cilindros_LB": cc.estandarizar_llave(Cilindros_LB),
                "Cilindros_KG": cc.estandarizar_llave(Cilindros_KG),
                "Cilindros_consol": cc.estandarizar_llave(Cilindros_consol),
                "resultados_tanques": cc.estandarizar_llave(resultados_tanques),
                "resultados_cilindros_tanques": cc.estandarizar_llave(resultado_CC_tanques),
                "Cap_Mer_Ini": cc.estandarizar_llave(Cap_Mer_Ini_consol),
                "Cap_Mer_Ope": cc.estandarizar_llave(Cap_Mer_Opera_cons),
                "Capacidad_Redes": cc.estandarizar_llave(Capacidad_Redes),
                "Capacidad_GLP_tot": cc.estandarizar_llave(Capacidad_GLP_tot),
                "participa_capacidad": cc.estandarizar_llave(participa_capacidad)
            }
          
        if file_Capac_anterior is not None:
            anterior_raw = pd.read_excel(file_Capac_anterior, sheet_name=None)

            anterior = {
                hoja: cc.estandarizar_llave(df)
                for hoja, df in anterior_raw.items()
            }

            comparaciones = cc.comparar_workbooks(actual, anterior, key="id_empresa", columnas_ignorar = ["Código SUI / Capacidad", "Empresa",
                             "Código SUI", "id_empresa", "EMPRESA", "ID_SUI", "RAZON_SOCIAL", "ID_EMPRESA", "Prestador", "valid_A", "valid_B", "valida_total",
                             "Código_SUI_num","EMPRES","indice_empresa"])

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
