            total_gasto = df["VALOR TOTAL"].sum()
            promedio = total_gasto / total_tiquetes if total_tiquetes else 0

            st.markdown('<div class="section-heading">Indicadores clave</div>', unsafe_allow_html=True)
            st.markdown('<div class="section-caption">Acumulado de los registros disponibles en la base de datos.</div>', unsafe_allow_html=True)
            k1, k2, k3, k4 = st.columns(4, gap="medium")
            with k1:
                metric_card("Comprobantes", f"{total_tiquetes:,}", "Registros consolidados")
            with k2:
                metric_card("Volumen total", f"{total_galones:,.1f} GAL", "Galones registrados")
            with k3:
                metric_card("Gasto acumulado", formato_cop(total_gasto), "Valor estimado de abastecimiento")
            with k4:
                metric_card("Promedio por registro", formato_cop(promedio), "Gasto medio por comprobante")

            st.markdown('<div class="section-heading">Lectura de la operación</div>', unsafe_allow_html=True)
            st.markdown('<div class="section-caption">Distribución del consumo y del gasto según la información registrada.</div>', unsafe_allow_html=True)
            chart1, chart2 = st.columns(2, gap="large")
            with chart1:
                df_matricula = (
                    df.groupby("MATRICULA", as_index=False)["CONSUMO (GAL)"].sum()
                    .sort_values("CONSUMO (GAL)", ascending=True)
                )
                fig_mat = px.bar(
                    df_matricula,
                    x="CONSUMO (GAL)",
                    y="MATRICULA",
                    orientation="h",
                    title="Consumo por aeronave",
                    labels={"CONSUMO (GAL)": "Galones", "MATRICULA": "Aeronave"},
                    color_discrete_sequence=[BLUE],
                )
                fig_mat.update_layout(
                    template="plotly_white", height=350,
                    margin=dict(l=18, r=20, t=65, b=25),
                    title=dict(font=dict(size=15, color=NAVY, family="Manrope")),
                    font=dict(family="DM Sans", color=MUTED, size=11),
                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                    xaxis=dict(showgrid=True, gridcolor="#EDF1F5", zeroline=False),
                    yaxis=dict(showgrid=False, zeroline=False),
                    showlegend=False,
                )
                fig_mat.update_traces(marker_line_width=0)
                st.plotly_chart(fig_mat, use_container_width=True, config={"displayModeBar": False})

            with chart2:
                df_ruta = df.groupby("RUTA", as_index=False)["VALOR TOTAL"].sum().sort_values("VALOR TOTAL", ascending=False)
                fig_ruta = px.pie(
                    df_ruta,
                    names="RUTA",
                    values="VALOR TOTAL",
                    title="Distribución del gasto por ruta",
                    hole=0.66,
                    color_discrete_sequence=[NAVY, BLUE, GOLD, "#8EA8C3", "#B9C9D9", "#6F879F"],
                )
                fig_ruta.update_layout(
                    template="plotly_white", height=350,
                    margin=dict(l=18, r=18, t=65, b=20),
                    title=dict(font=dict(size=15, color=NAVY, family="Manrope")),
                    font=dict(family="DM Sans", color=MUTED, size=11),
                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                    legend=dict(orientation="h", y=-0.12, x=0.5, xanchor="center", font=dict(size=10)),
                )
                fig_ruta.update_traces(textposition="inside", textinfo="percent", marker=dict(line=dict(color=WHITE, width=3)))
                st.plotly_chart(fig_ruta, use_container_width=True, config={"displayModeBar": False})

            st.markdown('<div class="section-heading">Registros de abastecimiento</div>', unsafe_allow_html=True)
            st.markdown('<div class="section-caption">Base maestra de comprobantes procesados.</div>', unsafe_allow_html=True)
            st.dataframe(
                df.sort_values("FECHA REGISTRO", ascending=False, na_position="last"),
                use_container_width=True,
                hide_index=True,
            )

            download_col, spacer_col = st.columns([1.1, 3.9])
            with download_col:
                with open(RUTA_EXCEL, "rb") as f:
                    st.download_button(
                        "Descargar base de datos",
                        data=f.read(),
                        file_name="Moonflight_Control_Combustible.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

            with st.expander("Administración y mantenimiento de datos"):
                st.warning("La eliminación es permanente. Descarga una copia de seguridad antes de continuar.")
                confirmar = st.checkbox("Confirmo que deseo eliminar todos los registros y soportes almacenados.")
                if st.button("Eliminar todos los registros", disabled=not confirmar):
                    if os.path.exists(RUTA_EXCEL):
                        os.remove(RUTA_EXCEL)
                    if os.path.exists(CARPETA_IMAGENES):
                        shutil.rmtree(CARPETA_IMAGENES)
                    inicializar_excel()
                    os.makedirs(CARPETA_IMAGENES, exist_ok=True)
                    st.success("Los registros fueron eliminados. Recarga el panel para ver la base vacía.")
                    st.rerun()
        else:
            st.markdown(
                """
                <div class="hero-panel">
                    <div class="hero-kicker">EXECUTIVE OVERVIEW</div>
                    <div class="hero-title">Tu centro de control está listo.</div>
                    <div class="hero-copy">Cuando se registren los primeros abastecimientos, aquí encontrarás los indicadores, las tendencias y la trazabilidad de los consumos de la flota.</div>
                    <div class="hero-mark">M.</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            e1, e2, e3 = st.columns(3, gap="medium")
            with e1:
                metric_card("Comprobantes", "0", "A la espera del primer registro")
            with e2:
                metric_card("Volumen total", "0 GAL", "Sin consumos registrados")
            with e3:
                metric_card("Gasto acumulado", "$ 0 COP", "Sin valores consolidados")
            st.info("Todavía no hay registros en la base de datos. Utiliza el módulo Registro en pista para comenzar.")

    st.markdown(
        '<div class="footer-line">MOONFLIGHT · FUEL INTELLIGENCE &nbsp; / &nbsp; EXECUTIVE DASHBOARD</div>',
        unsafe_allow_html=True,
    )
