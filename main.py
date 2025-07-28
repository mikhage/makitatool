import streamlit as st
import pandas as pd
from io import BytesIO
import matplotlib.pyplot as plt
import plotly.express as px
import plotly.graph_objects as go


st.set_page_config(page_title="CO₂ Calculator", layout="wide")
st.image("makita_logo.png", width=500)

# SESSION STATE
if "emission_factors" not in st.session_state:
    st.session_state.emission_factors = {"kWh": {"groen": 0.0, "grijs": 0.0}, "m³": {"groen": 0.0, "grijs": 0.0}}
if "elekfactor" not in st.session_state:
    st.session_state.elekfactor = 0.078
if "brandstof_factors" not in st.session_state:
    st.session_state.brandstof_factors = {"benzine": 0.0, "diesel": 0.0, "lpg": 0.0}
if "uploaded_file" not in st.session_state:
    st.session_state.uploaded_file = None
if "export_rows" not in st.session_state:
    st.session_state.export_rows = []
if "total_footprint" not in st.session_state:
    st.session_state.total_footprint = 0.0

page = st.sidebar.selectbox("Ga naar", ["Opties", "CO₂ Calculator", "Visualisaties"])
with open("user_manual.pdf", "rb") as file:
    st.sidebar.download_button(
        label="📄 Download Handleiding",
        data=file,
        file_name="user_manual.pdf",
        mime="application/pdf"
    )

# PAGINA: OPTIES
if page == "Opties":
    st.title("Instellingen: Emissiefactoren")
    excel_file = st.file_uploader("Upload Excel (koppen vanaf rij 5)", type=["xlsx"], key="factoren_uploader")
    if excel_file:
        df = pd.read_excel(excel_file, skiprows=4)
        df['normalized'] = df['Gasvormige brandstoffen'].astype(str).str.strip().str.lower()
        mapping = {
            ("kWh", "groen"): "windkracht",
            ("kWh", "grijs"): "grijze stroom",
            ("m³", "groen"): "groengas (gemiddeld)",
            ("m³", "grijs"): "aardgas (g-gas)",
        }
        for (unit, color), name in mapping.items():
            m = df['normalized'] == name
            if m.any():
                val = float(df.loc[m, 'Kg CO₂-eq / eenheid'].iloc[0])
                st.session_state.emission_factors[unit][color] = val
        bmap = {"benzine": "benzine (fossiel) e0", "diesel": "diesel (fossiel) b0", "lpg": "lpg"}
        for k, n in bmap.items():
            m = df['normalized'] == n
            if m.any():
                st.session_state.brandstof_factors[k] = float(df.loc[m, 'Kg CO₂-eq / eenheid'].iloc[0])
        st.success("Emissiefactoren automatisch geladen.")

    # Handmatig bijstellen
    st.subheader("Handmatig bijstellen (optioneel)")
    cols = st.columns([1, 1, 1])
    st.session_state.emission_factors['kWh']['groen'] = cols[0].number_input("kWh – Groen", value=st.session_state.emission_factors['kWh']['groen'], step=0.001, format="%.3f", key="kwh_g")
    st.session_state.emission_factors['kWh']['grijs'] = cols[1].number_input("kWh – Grijs", value=st.session_state.emission_factors['kWh']['grijs'], step=0.001, format="%.3f", key="kwh_r")
    st.session_state.elekfactor = cols[2].number_input("Elektrisch vervoer (kg CO₂/kWh)", value=st.session_state.elekfactor, step=0.001, format="%.3f", key="elek")
    st.session_state.emission_factors['m³']['groen'] = cols[0].number_input("m³ – Groen", value=st.session_state.emission_factors['m³']['groen'], step=0.001, format="%.3f", key="m3_g")
    st.session_state.emission_factors['m³']['grijs'] = cols[1].number_input("m³ – Grijs", value=st.session_state.emission_factors['m³']['grijs'], step=0.001, format="%.3f", key="m3_r")

    for b in ['benzine', 'diesel', 'lpg']:
        st.session_state.brandstof_factors[b] = st.number_input(f"{b.capitalize()} (kg CO₂/L)", value=st.session_state.brandstof_factors[b], step=0.001, format="%.3f", key=f"bf_{b}")

# PAGINA: CO₂ Calculator
elif page == "CO₂ Calculator":
    st.title("CO₂ Calculator")
    f = st.file_uploader("Upload Excel met tabbladen", type=["xlsx"])
    if f and f != st.session_state.uploaded_file:
        st.session_state.uploaded_file = f
        st.session_state.export_rows = []
        st.session_state.total_footprint = 0.0

    if st.session_state.uploaded_file:
        xl = pd.ExcelFile(st.session_state.uploaded_file)
        st.markdown("### Invoeroverzicht")
        hdr = st.columns([3, 1, 3, 2, 2])
        hdr[0].markdown("**Onderdeel**")
        hdr[1].markdown("**Eenheid**")
        hdr[2].markdown("**Emissiefactor**")
        hdr[3].markdown("**Verbruik**")
        hdr[4].markdown("**Footprint**")

        total_fp = 0.0
        rows = []

        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            if 'Brandstof' in df.columns:
                for bt in df['Brandstof'].dropna().unique():
                    sub = df[df['Brandstof'] == bt]
                    verbruik = pd.to_numeric(sub['Brandstof p/j'], errors='coerce').sum()
                    bt_lc = bt.lower()
                    if any(k in bt_lc for k in ['elektrisch', 'elektriciteit', 'ev']):
                        fact = st.session_state.elekfactor
                        een = 'kWh'
                        name = f"Vervoer-E"
                    elif 'hybride' in bt_lc:
                        fact = st.session_state.brandstof_factors.get('benzine', 0.0)
                        een = 'L'
                        name = f"Vervoer-Hybride"
                    else:
                        een = 'L'
                        key = bt_lc.split()[0]
                        fact = st.session_state.brandstof_factors.get(key, 0.0)
                        name = f"Vervoer-{bt}"
                    fp = verbruik * fact
                    total_fp += fp
                    rows.append({'Onderdeel': name, 'Eenheid': een, 'Emissiefactor': fact, 'Verbruik': verbruik, 'Footprint': fp})
                    cols = st.columns([3, 1, 3, 2, 2])
                    cols[0].markdown(name)
                    cols[1].markdown(een)
                    cols[2].markdown(f"**{fact:.3f}**")
                    cols[3].markdown(f"{verbruik:,.2f}")
                    cols[4].markdown(f"**{fp:,.2f}**")

            elif all(c in df.columns for c in ['Aantal', 'Vermogen', 'Eenheid', 'Draaiuren p/j']):
                df['Verbruik'] = pd.to_numeric(df['Aantal'], errors='coerce') * pd.to_numeric(df['Vermogen'], errors='coerce') * pd.to_numeric(df['Draaiuren p/j'], errors='coerce')
                verbruik = df['Verbruik'].sum()
                een = df['Eenheid'].dropna().iloc[0]
                if een in st.session_state.emission_factors:
                    c1, c2, c3, c4, c5 = st.columns([3, 1, 3, 2, 2])
                    k = c3.radio("Stroomtype", ['Groen', 'Grijs'], horizontal=True, key=sheet + een)
                    fact = st.session_state.emission_factors[een][k.lower()]
                elif een == 'L':
                    fuel = st.selectbox(f"Brandstof '{sheet}'", options=list(st.session_state.brandstof_factors.keys()), key='f' + sheet)
                    fact = st.session_state.brandstof_factors[fuel]
                else:
                    fact = 0.0
                fp = verbruik * fact
                total_fp += fp
                rows.append({'Onderdeel': sheet, 'Eenheid': een, 'Emissiefactor': fact, 'Verbruik': verbruik, 'Footprint': fp})
                c = st.columns([3, 1, 3, 2, 2])
                c[0].markdown(sheet)
                c[1].markdown(een)
                c[2].markdown(f"**{fact:.3f}**")
                c[3].markdown(f"{verbruik:,.2f}")
                c[4].markdown(f"**{fp:,.2f}**")

        st.session_state.total_footprint = total_fp
        st.session_state.export_rows = rows
        st.markdown("---")
        st.subheader(f"Totale CO₂-footprint: **{total_fp:,.2f} kg CO₂**")

        dfout = pd.DataFrame(rows)
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            dfout.to_excel(writer, index=False, sheet_name='Resultaten')
        buf.seek(0)
        st.download_button("📥 Download resultaten", data=buf, file_name="CO2_resultaten.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("Upload een Excelbestand met de juiste kolommen.")

elif page == "Visualisaties":
    st.title("Visualisaties")
    
    # Tabs voor verschillende visualisatie opties
    tab1, tab2 = st.tabs(["📊 Per Tabblad", "📈 CO₂-Intensiteit Trend"])
    
    with tab1:
        st.subheader("Visualisaties per Excel-tabblad")
        f = st.session_state.uploaded_file
        if f:
            xl = pd.ExcelFile(f)

            tabblad_verbruiken = {}
            tabblad_footprints = {}

            for sheet in xl.sheet_names:
                st.subheader(f"Tabblad: {sheet}")
                df = xl.parse(sheet)

                # ——— VOOR VERVOER TABBLADEN ———
                if 'vervoer' in sheet.lower() and all(col in df.columns for col in ['Brandstof', 'Merk', 'Type', 'Brandstof p/j']):
                    df['Brandstof p/j'] = pd.to_numeric(df['Brandstof p/j'], errors='coerce')
                    df = df.dropna(subset=['Brandstof p/j'])
                    top5 = df.sort_values(by='Brandstof p/j', ascending=False).head(5)

                    st.markdown("#### Top 5 hoogste brandstofverbruik")
                    st.dataframe(
                        top5[['Merk', 'Type', 'Brandstof', 'Brandstof p/j']].rename(
                            columns={'Brandstof p/j': 'Verbruik (L)'}
                        ),
                        use_container_width=True
                    )

                    # Totaal verbruik en CO2-footprint voor taartdiagrammen
                    totaal_verbruik = df['Brandstof p/j'].sum()
                    tabblad_verbruiken[sheet] = totaal_verbruik

                    # Bepaal factor op basis van eerste rij (simpele benadering)
                    brandstof = df['Brandstof'].dropna().iloc[0].lower() if not df['Brandstof'].dropna().empty else ''
                    factor = st.session_state.brandstof_factors.get(brandstof, 0.0)
                    footprint = totaal_verbruik * factor
                    tabblad_footprints[sheet] = footprint

                # ——— VOOR ANDERE TABBLADEN MET APPARATEN ———
                elif all(col in df.columns for col in ['Aantal', 'Vermogen', 'Draaiuren p/j']):
                    df['Aantal'] = pd.to_numeric(df['Aantal'], errors='coerce')
                    df['Vermogen'] = pd.to_numeric(df['Vermogen'], errors='coerce')
                    df['Draaiuren p/j'] = pd.to_numeric(df['Draaiuren p/j'], errors='coerce')
                    df = df.dropna(subset=['Aantal', 'Vermogen', 'Draaiuren p/j'])
                    df = df[df['Aantal'] > 0]

                    df['Verbruik'] = (df['Vermogen'] * df['Draaiuren p/j']) / df['Aantal']
                    df['Type'] = df.get('Type', 'Onbekend')
                    df['Merk'] = df.get('Merk', 'Onbekend')

                    top5 = df.sort_values(by='Verbruik', ascending=False).head(5)

                    st.markdown("#### Top 5 hoogste verbruik per apparaat")
                    st.dataframe(
                        top5[['Merk', 'Type', 'Verbruik']].rename(columns={
                            'Verbruik': 'Verbruik per apparaat'
                        }),
                        use_container_width=True
                    )

                    # Totaal verbruik = vermogen * draaiuren (over alles)
                    totaal_verbruik = (df['Vermogen'] * df['Draaiuren p/j']).sum()
                    tabblad_verbruiken[sheet] = totaal_verbruik

                    # Emissiefactor bepalen op basis van eerste 'Eenheid'
                    eenheid = df['Eenheid'].dropna().iloc[0] if 'Eenheid' in df.columns and not df['Eenheid'].dropna().empty else None
                    if eenheid and eenheid in st.session_state.emission_factors:
                        factor = st.session_state.emission_factors[eenheid]['grijs']  # bijv. 'kWh'
                    else:
                        factor = 0.0

                    footprint = totaal_verbruik * factor
                    tabblad_footprints[sheet] = footprint

                else:
                    st.warning("Tabblad heeft geen herkenbare structuur.")

            # ——— TOP 5 VERBRUIK PER TABBLAD ———
            if tabblad_verbruiken:
                st.markdown("#### Top 5 hoogste totaalverbruik per tabblad")
                # Maak een DataFrame voor de tabblad verbruiken en sorteer op verbruik
                top5_verbruik_df = pd.DataFrame({
                    'Tabblad': list(tabblad_verbruiken.keys()),
                    'Totaal verbruik': list(tabblad_verbruiken.values())
                }).sort_values(by='Totaal verbruik', ascending=False).head(5)

                st.dataframe(top5_verbruik_df, use_container_width=True)

            # ——— TOP 5 CO2-FOOTPRINT PER TABBLAD ———
            if tabblad_footprints:
                st.markdown("#### Top 5 hoogste CO₂-footprint per tabblad")
                # Maak een DataFrame voor de tabblad footprints en sorteer op CO₂-footprint
                top5_footprint_df = pd.DataFrame({
                    'Tabblad': list(tabblad_footprints.keys()),
                    'CO2-footprint (kg)': list(tabblad_footprints.values())
                }).sort_values(by='CO2-footprint (kg)', ascending=False).head(5)

                st.dataframe(top5_footprint_df, use_container_width=True)

            # ——— CIRKELDIAGRAM TOTAALVERBRUIK PER TABBLAD ———
            if tabblad_verbruiken:
                st.markdown("### Verhouding totaal verbruik per tabblad")
                chart_df = pd.DataFrame({
                    'Tabblad': list(tabblad_verbruiken.keys()),
                    'Totaal verbruik': list(tabblad_verbruiken.values())
                })
                fig = px.pie(chart_df, names='Tabblad', values='Totaal verbruik', hole=0.4)
                st.plotly_chart(fig, use_container_width=True)

            # ——— CIRKELDIAGRAM CO2-FOOTPRINT PER TABBLAD ———
            if tabblad_footprints:
                st.markdown("### Verhouding totale CO₂-footprint per tabblad")
                footprint_df = pd.DataFrame({
                    'Tabblad': list(tabblad_footprints.keys()),
                    'CO2-footprint (kg)': list(tabblad_footprints.values())
                })
                fig = px.pie(footprint_df, names='Tabblad', values='CO2-footprint (kg)', hole=0.4)
                st.plotly_chart(fig, use_container_width=True)

        else:
            st.info("Upload eerst een Excelbestand via de CO₂ Calculator-pagina.")
    
    with tab2:
        st.subheader("CO₂-Intensiteit Trend Analyse")
        st.markdown("Upload een Excel-bestand met de volgende kolommen: **Jaar**, **Omzet (miljoenen)**, **Co2-Footprint (ton)**")
        
        # Upload voor CO2-intensiteit data
        intensity_file = st.file_uploader(
            "Upload Excel voor CO₂-intensiteit analyse", 
            type=["xlsx"], 
            key="intensity_uploader"
        )
        
        if intensity_file:
            try:
                # Lees het Excel bestand
                df_intensity = pd.read_excel(intensity_file)
                
                # Verwachte kolomnamen (flexibel voor variaties met veel substrings)
                jaar_substrings = ['jaar', 'year', 'periode', 'datum', 'date', 'time', 'tijd', 'boekjaar']
                omzet_substrings = ['omzet', 'revenue', 'turnover', 'sales', 'verkoop', 'inkomsten', 'opbrengst', 'netto', 'bruto', 'facturatie', 'totaal']
                co2_substrings = ['co2', 'co₂', 'footprint', 'uitstoot', 'emissie', 'carbon', 'koolstof', 'milieu', 'duurzaam', 'klimaat', 'scope', 'ghg', 'greenhouse']
                
                jaar_cols = [col for col in df_intensity.columns if any(substring in col.lower() for substring in jaar_substrings)]
                omzet_cols = [col for col in df_intensity.columns if any(substring in col.lower() for substring in omzet_substrings)]
                co2_cols = [col for col in df_intensity.columns if any(substring in col.lower() for substring in co2_substrings)]
                
                if not jaar_cols or not omzet_cols or not co2_cols:
                    st.error("⚠️ Controleer of het bestand de juiste kolommen heeft: Jaar, Omzet (miljoenen), Co2-Footprint (ton)")
                    st.write("Gevonden kolommen:", list(df_intensity.columns))
                else:
                    # Gebruik de eerste gevonden kolom van elk type
                    jaar_col = jaar_cols[0]
                    omzet_col = omzet_cols[0]
                    co2_col = co2_cols[0]
                    
                    # Data cleanup
                    df_clean = df_intensity[[jaar_col, omzet_col, co2_col]].copy()
                    df_clean.columns = ['Jaar', 'Omzet_miljoen', 'CO2_ton']
                    
                    # Converteer naar numeriek
                    df_clean['Jaar'] = pd.to_numeric(df_clean['Jaar'], errors='coerce')
                    df_clean['Omzet_miljoen'] = pd.to_numeric(df_clean['Omzet_miljoen'], errors='coerce')
                    df_clean['CO2_ton'] = pd.to_numeric(df_clean['CO2_ton'], errors='coerce')
                    
                    # Verwijder rijen met NaN waarden
                    df_clean = df_clean.dropna()
                    
                    if df_clean.empty:
                        st.error("Geen geldige data gevonden na het opschonen.")
                    else:
                        # Bereken CO2-intensiteit (ton CO2 per miljoen euro omzet)
                        df_clean['CO2_Intensiteit'] = df_clean['CO2_ton'] / df_clean['Omzet_miljoen']
                        
                        # Sorteer op jaar
                        df_clean = df_clean.sort_values('Jaar')
                        
                        # Toon de data
                        st.markdown("### Gegevensoverzicht")
                        display_df = df_clean.copy()
                        display_df['CO2_Intensiteit'] = display_df['CO2_Intensiteit'].round(3)
                        display_df.columns = ['Jaar', 'Omzet (miljoen €)', 'CO₂-Footprint (ton)', 'CO₂-Intensiteit (ton/miljoen €)']
                        st.dataframe(display_df, use_container_width=True)
                        
                        # Visualisaties
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.markdown("### CO₂-Intensiteit Trend")
                            fig_trend = go.Figure()
                            fig_trend.add_trace(go.Scatter(
                                x=df_clean['Jaar'],
                                y=df_clean['CO2_Intensiteit'],
                                mode='lines+markers',
                                name='CO₂-Intensiteit',
                                line=dict(color='#FF6B6B', width=3),
                                marker=dict(size=8)
                            ))
                            fig_trend.update_layout(
                                title="CO₂-Intensiteit over de jaren",
                                xaxis_title="Jaar",
                                yaxis_title="CO₂-Intensiteit (ton CO₂/miljoen €)",
                                hovermode='x unified'
                            )
                            st.plotly_chart(fig_trend, use_container_width=True)
                            
                        with col2:
                            st.markdown("### Omzet vs CO₂-Footprint")
                            fig_scatter = go.Figure()
                            fig_scatter.add_trace(go.Scatter(
                                x=df_clean['Omzet_miljoen'],
                                y=df_clean['CO2_ton'],
                                mode='markers+text',
                                text=df_clean['Jaar'].astype(str),
                                textposition="top center",
                                name='Jaardata',
                                marker=dict(size=12, color='#4ECDC4')
                            ))
                            fig_scatter.update_layout(
                                title="Omzet vs CO₂-Footprint per jaar",
                                xaxis_title="Omzet (miljoen €)",
                                yaxis_title="CO₂-Footprint (ton)"
                            )
                            st.plotly_chart(fig_scatter, use_container_width=True)
                        
                        # Analyse samenvatting
                        st.markdown("### Analyse Samenvatting")
                        
                        if len(df_clean) > 1:
                            start_intensiteit = df_clean.iloc[0]['CO2_Intensiteit']
                            eind_intensiteit = df_clean.iloc[-1]['CO2_Intensiteit']
                            verandering_pct = ((eind_intensiteit - start_intensiteit) / start_intensiteit) * 100
                            
                            col1, col2, col3, col4 = st.columns(4)
                            
                            with col1:
                                st.metric(
                                    "Huidige CO₂-Intensiteit", 
                                    f"{eind_intensiteit:.3f} ton/M€"
                                )
                            
                            with col2:
                                st.metric(
                                    "Startwaarde", 
                                    f"{start_intensiteit:.3f} ton/M€"
                                )
                            
                            with col3:
                                st.metric(
                                    "Verandering (%)", 
                                    f"{verandering_pct:+.1f}%",
                                    delta=f"{verandering_pct:+.1f}%"
                                )
                            
                            with col4:
                                beste_jaar = df_clean.loc[df_clean['CO2_Intensiteit'].idxmin(), 'Jaar']
                                beste_waarde = df_clean['CO2_Intensiteit'].min()
                                st.metric(
                                    f"Beste jaar ({int(beste_jaar)})", 
                                    f"{beste_waarde:.3f} ton/M€"
                                )
                            
                            # Interpretatie
                            st.markdown("### Interpretatie")
                            if verandering_pct < -5:
                                st.success("✅ **Uitstekend!** Je CO₂-intensiteit is significant gedaald. Dit betekent dat je bedrijf veel efficiënter is geworden.")
                            elif verandering_pct < 0:
                                st.success("✅ **Goed!** Je CO₂-intensiteit is gedaald. Je bedrijf wordt groener.")
                            elif verandering_pct < 5:
                                st.warning("⚠️ **Stabiel.** Je CO₂-intensiteit is relatief stabiel gebleven.")
                            else:
                                st.error("❌ **Aandacht vereist.** Je CO₂-intensiteit is gestegen. Overweeg duurzaamheidsmaatregelen.")
                        
                        # Download optie
                        buf = BytesIO()
                        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                            display_df.to_excel(writer, index=False, sheet_name='CO2_Intensiteit_Analyse')
                        buf.seek(0)
                        
                        st.download_button(
                            "📥 Download CO₂-intensiteit analyse",
                            data=buf,
                            file_name="CO2_intensiteit_analyse.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
            except Exception as e:
                st.error(f"Fout bij het verwerken van het bestand: {str(e)}")
                st.info("Zorg ervoor dat het Excel-bestand de juiste kolommen heeft: Jaar, Omzet (miljoenen), Co2-Footprint (ton)")
        
        else:
            st.info("📁 Upload een Excel-bestand om de CO₂-intensiteit trend te analyseren.")