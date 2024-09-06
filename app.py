import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.express as px
from datetime import datetime

# Ajouter le logo à la barre latérale avec un style CSS pour le déplacer légèrement vers la gauche
st.sidebar.markdown(
    """
    <style>
        .sidebar-content img {
            margin-left: -50px;  /* Ajustez cette valeur pour déplacer le logo vers la gauche */
        }
    </style>
    """,
    unsafe_allow_html=True
)

# Ajouter le logo à la barre latérale
st.sidebar.image("logo.png", width=380)

# Ajouter un titre à la barre latérale
st.sidebar.markdown("# Tableau de bord")

# Sélectionner le type de données à afficher
data_type = st.sidebar.selectbox("Sélectionnez le type de données", ["Finance", "Ressources humaines"])

# Créer deux colonnes pour afficher le logo
empty_col, logo_col = st.columns([1, 4])  # Ajustez les proportions pour contrôler l'espace

# Colonne vide (gauche)
with empty_col:
    st.write("")  # Colonne vide pour créer un espace

# Charger le fichier depuis la barre latérale
uploaded_file = st.sidebar.file_uploader("Choisir un fichier", type=["xlsx", "xls", "csv"])

if uploaded_file is None:
    # Afficher le logo uniquement lorsque le fichier n'est pas encore téléchargé
    col1, col2 = st.columns([1, 4])  # Ajustez les proportions pour contrôler l'espace

    # Colonne vide (gauche)
    with col1:
        st.write("")  # Colonne vide pour créer un espace

    # Colonne pour afficher le logo (droite)
    with col2:
        st.image("Lo.png", width=400)  # Ajustez la largeur du logo selon vos besoins

    # Paragraphe combiné avec un léger décalage vers le haut
    st.markdown(
        """
        <div style="solid #dcdcdc; border-radius: 10px; padding: 20px; margin-top: -70px;">
            <p style="font-size: 20px; text-align: justify;">
            <span style="color: #003366;"><strong>Bienvenue sur Capital & Talent View</strong></span>, votre solution tout-en-un pour la visualisation et l'analyse de données financières et des ressources humaines. 
            Cette application a été conçue pour vous offrir une vue claire et interactive de vos données, 
            facilitant ainsi la prise de décision éclairée.</p>
            <p style="font-size: 20px; text-align: justify;">Avec Capital & Talent View, vous pouvez facilement explorer vos données financières et RH, 
            visualiser les tendances sur plusieurs années, et comparer les indicateurs clés de performance à travers différents programmes et projets. 
            Chargez simplement vos fichiers Excel et laissez l'application faire le reste. Que vous soyez gestionnaire, analyste, ou cadre, 
            Capital & Talent View vous aide à transformer vos données en informations exploitables.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

def calculate_indicators(df):
    # Convertir les colonnes en numériques
    df['Crédits ouverts  (A)'] = pd.to_numeric(df['Crédits ouverts  (A)'], errors='coerce')
    df['Engagements \n(B)'] = pd.to_numeric(df['Engagements \n(B)'], errors='coerce')
    df['Paiements \n(C)'] = pd.to_numeric(df['Paiements \n(C)'], errors='coerce')
    
    total_credits = df['Crédits ouverts  (A)'].sum()
    total_engagements = df['Engagements \n(B)'].sum()
    total_payments = df['Paiements \n(C)'].sum()
    return total_credits, total_engagements, total_payments

# Assurez-vous que 'uploaded_file' est défini dans votre code avant d'utiliser cette condition
if data_type == "Finance":
    try:
        # Lecture des données financières
        excel_data = pd.ExcelFile(uploaded_file)
        # Ajouter la nouvelle visualisation pour les sources de budget
        
       # Charger les données du fichier Excel pour le budget
        budget_2022 = pd.read_excel(excel_data, sheet_name='Budget_CNRST_Source_2022')
        budget_2023 = pd.read_excel(excel_data, sheet_name='Budget_CNRST_Source_2023')
        # Convertir les colonnes numériques
        budget_2022['Dotation reçue'] = pd.to_numeric(budget_2022['Dotation reçue'], errors='coerce').fillna(0)
        budget_2023['Dotation reçue'] = pd.to_numeric(budget_2023['Dotation reçue'], errors='coerce').fillna(0)


        # Ajouter les colonnes 'Year'
        budget_2022['Year'] = 2022
        budget_2023['Year'] = 2023

        # Combiner les données en un seul DataFrame
        budget_all_years = pd.concat([budget_2022, budget_2023])

        # Filtrer les données pour les sources spécifiques
        selected_sources = ["Budget général d'Etat (ministère de tutelle)", "FNSRSDT", "Recettes propres"]
        filtered_data = budget_all_years[budget_all_years['Source'].isin(selected_sources)]

        # Calcul des indicateurs pour le budget
        total_budget_2022 = budget_2022['Dotation reçue'].sum()
        total_budget_2023 = budget_2023['Dotation reçue'].sum()

        # Afficher les indicateurs de budget dans des colonnes
        col1, col2 = st.columns(2)

        with col1:
            st.info("Budget Total 2022", icon="📊")
            st.metric(label="Total Dotation 2022", value=f"{total_budget_2022:.2f}")

        with col2:
            st.info("Budget Total 2023", icon="💰")
            st.metric(label="Total Dotation 2023", value=f"{total_budget_2023:.2f}")

        # Créer la figure pour le graphique en lignes
        fig_sources = go.Figure()

        # Définir les couleurs pour chaque source
        source_colors = {
            "Budget général d'Etat (ministère de tutelle)": 'blue',
            "FNSRSDT": 'green',
            "Recettes propres": 'red'
        }

        # Ajouter une ligne pour chaque source sélectionnée avec une couleur spécifique
        for source in selected_sources:
            df_source = filtered_data[filtered_data['Source'] == source]
            fig_sources.add_trace(go.Scatter(
                x=df_source['Year'],
                y=df_source['Dotation reçue'],
                mode='lines+markers',  # Ajouter des marqueurs à chaque point de données
                name=source,
                line=dict(color=source_colors[source], width=2)
            ))

        # Mettre à jour la mise en page
        fig_sources.update_layout(
            title="Evolution de Budget Sources",
            xaxis_title="Year",
            yaxis_title="Dotation reçue",
            legend_title="Source",
            xaxis=dict(tickmode='linear'),  
            )
        # Affichage du graphique de lignes dans Streamlit
        st.plotly_chart(fig_sources)


        # Traitement des données d'exploitation
        exploitation_2021 = pd.read_excel(excel_data, sheet_name='Exploitation 2021')
        exploitation_2022 = pd.read_excel(excel_data, sheet_name='Exploitation 2022')

        # Ajout des colonnes 'Year' pour exploitation
        exploitation_2021['Year'] = 2021
        exploitation_2022['Year'] = 2022

        # Combinaison des données d'exploitation
        exploitation_all_years = pd.concat([exploitation_2021, exploitation_2022])
        # Création de la figure pour exploitation
        fig_exploitation = make_subplots(rows=1, cols=1, shared_xaxes=True)

        # Ajout des traces pour chaque année (exploitation)
        for year in [2021, 2022]:
            df_year = exploitation_all_years[exploitation_all_years['Year'] == year]
            for financial_type in ['Crédits ouverts  (A)', 'Engagements \n(B)', 'Paiements \n(C)']:
                fig_exploitation.add_trace(go.Bar(
                    x=df_year['Programme'],
                    y=df_year[financial_type],
                    name=f'{financial_type} - {year}',
                    visible='legendonly'
                ))
                # Obtention de l'année actuelle
                current_year = datetime.now().year
        # Ajout du menu déroulant pour exploitation
        fig_exploitation.update_layout(
            updatemenus=[
                dict(
                    type="dropdown",
                    showactive=True,
                    buttons=[
                        dict(label="2021",
                             method="update",
                             args=[{"visible": [True if '2021' in trace.name else False for trace in fig_exploitation.data]},
                                   {"title": "Exploitation - 2021"}]),
                        dict(label="2022",
                             method="update",
                             args=[{"visible": [True if '2022' in trace.name else False for trace in fig_exploitation.data]},
                                   {"title": "Exploitation - 2022"}]),
                        dict(label=str(current_year),
                         method="update",
                         args=[{"visible": [True if f'{current_year}' in trace.name else False for trace in fig_exploitation.data]},
                           {"title": f"Investment Data - {current_year}"}])
                    ]
                )
            ],
            title="Financial Data by Year",
            xaxis_title="Programme",
            yaxis_title="Value"
        )

        # Affichage de la figure exploitation dans Streamlit
        st.plotly_chart(fig_exploitation)

        # Lecture des données d'investissement
        investment_2021 = pd.read_excel(excel_data, sheet_name='Invistissement2021')
        investment_2022 = pd.read_excel(excel_data, sheet_name='Investissement2022')
        investment_2023 = pd.read_excel(excel_data, sheet_name='Investissement 2023')

        # Ajout des colonnes 'Year' pour investissement
        investment_2021['Year'] = 2021
        investment_2022['Year'] = 2022
        investment_2023['Year'] = 2023

        # Combinaison des données d'investissement
        investment_all_years = pd.concat([investment_2021, investment_2022, investment_2023])

        # Sélection des colonnes à afficher pour investissement
        columns_to_plot_investment = ['Crédits ouverts Hors reports (A)', 'Engagements Hors Reports (B)', '% Engagement',
                                      'Paiements Hors Reports (C)', '% Paiement', 'Reports\nD= (B-C)', 'Disponible (A-B)']

        # Création de la figure pour investissement
        fig_investment = make_subplots(rows=1, cols=1, shared_xaxes=True)

        # Ajout des traces pour chaque année (investissement)
        for year in [2021, 2022, 2023]:
            df_year = investment_all_years[investment_all_years['Year'] == year]
            for financial_type in columns_to_plot_investment:
                fig_investment.add_trace(go.Bar(
                    x=df_year['Programme'],
                    y=df_year[financial_type],
                    name=f'{financial_type} - {year}',
                    visible='legendonly'
                ))
                # Obtention de l'année actuelle
            current_year = datetime.now().year

         #Ajout du menu déroulant pour investissement
        fig_investment.update_layout(
            updatemenus=[
        dict(
            type="dropdown",
            showactive=True,
            buttons=[
                dict(label="All Years",
                     method="update",
                     args=[{"visible": [True] * len(fig_investment.data)},
                           {"title": "Investment Data by Year"}]),
                dict(label="2021",
                     method="update",
                     args=[{"visible": [True if '2021' in trace.name else False for trace in fig_investment.data]},
                           {"title": "Investment Data - 2021"}]),
                dict(label="2022",
                     method="update",
                     args=[{"visible": [True if '2022' in trace.name else False for trace in fig_investment.data]},
                           {"title": "Investment Data - 2022"}]),
                dict(label="2023",
                     method="update",
                     args=[{"visible": [True if '2023' in trace.name else False for trace in fig_investment.data]},
                           {"title": "Investment Data - 2023"}]),
                dict(label=str(current_year),
                     method="update",
                     args=[{"visible": [True if f'{current_year}' in trace.name else False for trace in fig_investment.data]},
                           {"title": f"Investment Data - {current_year}"}])
            ]
        )
    ],
    title="Investment Data by Year",
    xaxis_title="Programme",
    yaxis_title="Value"
)
        # Affichage de la figure investissement dans Streamlit
        st.plotly_chart(fig_investment)

       
         # Load data from 'Marchés Pluriannuel' sheets
        marches_2021 = pd.read_excel(excel_data, sheet_name='Marchés Pluriannuel 2021')
        marches_2022 = pd.read_excel(excel_data, sheet_name='Marchés Pluriannuel 2022')
        marches_2023 = pd.read_excel(excel_data, sheet_name='Marchés Pluriannuel 2023')

        # Add 'Year' columns
        marches_2021['Year'] = 2021
        marches_2022['Year'] = 2022
        marches_2023['Year'] = 2023

        # Combine the data into one DataFrame
        marches_all_years = pd.concat([marches_2021, marches_2022, marches_2023])

        # Define the columns to plot
        columns_to_plot = ['Nombre de contrats', 'Montant en Dhs ']

        # Create the bar chart figure
        fig_marches = go.Figure()

        # Add traces for each type of data for each year
        for year in [2021, 2022, 2023]:
            df_year = marches_all_years[marches_all_years['Year'] == year]
            for column in columns_to_plot:
                fig_marches.add_trace(go.Bar(
                    x=df_year['Objet'],
                    y=df_year[column],
                    name=f'{column} - {year}',
                    visible='legendonly'
                ))

        # Add dropdown menu for year selection
        fig_marches.update_layout(
            updatemenus=[
                dict(
                    type="dropdown",
                    showactive=True,
                    buttons=[
                        dict(label="2021",
                             method="update",
                             args=[{"visible": [True if '2021' in trace.name else False for trace in fig_marches.data]},
                                   {"title": "Marchés Pluriannuel Data - 2021"}]),
                        dict(label="2022",
                             method="update",
                             args=[{"visible": [True if '2022' in trace.name else False for trace in fig_marches.data]},
                                   {"title": "Marchés Pluriannuel Data - 2022"}]),
                        dict(label="2023",
                             method="update",
                             args=[{"visible": [True if '2023' in trace.name else False for trace in fig_marches.data]},
                                   {"title": "Marchés Pluriannuel Data - 2023"}]),
                        dict(label=str(current_year),
                            method="update",
                            args=[{"visible": [True if f'{current_year}' in trace.name else False for trace in fig_marches.data]},
                           {"title": f"Investment Data - {current_year}"}])
                    ]
                )
            ],
            title="Marchés Pluriannuel Data by Year",
            xaxis_title="Objet",
            yaxis_title="Value",
            barmode='group'
        )

        # Display the interactive bar chart in Streamlit
        st.plotly_chart(fig_marches)
          
      # New Visualization: Achat Marché Nature Exploit by Year
        nature_exploit_2022 = pd.read_excel(excel_data, sheet_name='Achat Marché nature_Exploit2022')
        nature_exploit_2023 = pd.read_excel(excel_data, sheet_name='Achat Marché nature_Exploit2023')

        # Add 'Year' columns
        nature_exploit_2022['Year'] = 2022
        nature_exploit_2023['Year'] = 2023

        # Combine the data into one dataframe
        nature_combined = pd.concat([nature_exploit_2022, nature_exploit_2023])

        # Create the bar chart figure
        fig_nature = go.Figure()

        # Add traces for each type of data for each year
        for year in [2022, 2023]:
            df_year = nature_combined[nature_combined['Year'] == year]
            fig_nature.add_trace(go.Bar(
                x=df_year['Nature'],
                y=df_year['Nombre de marché '],
                name=f'Nombre de marché - {year}',
                visible=(year == 2022)  # Only the first year is visible initially
            ))
            fig_nature.add_trace(go.Bar(
                x=df_year['Nature'],
                y=df_year['Montant en Dhs'],
                name=f'Montant en Dhs - {year}',
                visible=(year == 2022)  # Only the first year is visible initially
            ))

        # Add dropdown menu for year selection
        fig_nature.update_layout(
            updatemenus=[
                dict(
                    type="dropdown",
                    showactive=True,
                    buttons=[
                        dict(label="2022",
                             method="update",
                             args=[{"visible": [True if '2022' in trace.name else False for trace in fig_nature.data]},
                                   {"title": "Achat Marché Nature Exploit 2022"}]),
                        dict(label="2023",
                             method="update",
                             args=[{"visible": [True if '2023' in trace.name else False for trace in fig_nature.data]},
                                   {"title": "Achat Marché Nature Exploit 2023"}]),
                        dict(label=str(current_year),
                            method="update",
                         args=[{"visible": [True if f'{current_year}' in trace.name else False for trace in fig_nature.data]},
                           {"title": f"Investment Data - {current_year}"}])
                    ]
                )
            ],
            title="Achat Marché Nature Exploit by Year",
            xaxis_title="Nature",
            yaxis_title="Value",
            barmode='group',
            legend_title="Legend",
            xaxis_tickangle=-45
        )

         # Display the interactive bar chart in Streamlit
        st.plotly_chart(fig_nature)
          
          # Load data from the sheets
        achat_nature_2021 = pd.read_excel(excel_data, sheet_name='Achat Marché  nature Invis2021')
        achat_nature_2022 = pd.read_excel(excel_data, sheet_name='Achat Marché nature Invisti2022')
        achat_nature_2023 = pd.read_excel(excel_data, sheet_name='Achat Marché par nature Inves23')

         # Add 'Year' columns
        achat_nature_2021['Year'] = 2021
        achat_nature_2022['Year'] = 2022
        achat_nature_2023['Year'] = 2023

        # Combine the data into one dataframe
        achat_nature_all_years = pd.concat([achat_nature_2021, achat_nature_2022, achat_nature_2023])

        # Define the columns to plot
        columns_to_plot = ['Nombre de marché ', 'Montant en Dhs']

        # Create the stacked bar chart figure
        fig = go.Figure()

         # Add traces for each type of data for each year
        for year in [2021, 2022, 2023]:
           df_year = achat_nature_all_years[achat_nature_all_years['Year'] == year]
           fig.add_trace(go.Bar(
           x=df_year['Nature'],
           y=df_year['Nombre de marché '],
           name=f'Nombre de marché - {year}',
           visible='legendonly'
        ))
           fig.add_trace(go.Bar(
           x=df_year['Nature'],
            y=df_year['Montant en Dhs'],
            name=f'Montant en Dhs - {year}',
            visible='legendonly'
        ))

         # Add dropdown menu for year selection
        fig.update_layout(
              updatemenus=[
        dict(
            type="dropdown",
            showactive=True,
            buttons=[
                dict(label="2021",
                     method="update",
                     args=[{"visible": [True if '2021' in trace.name else False for trace in fig.data]},
                           {"title": "Achat Marché Nature investissemnt- 2021"}]),
                dict(label="2022",
                     method="update",
                     args=[{"visible": [True if '2022' in trace.name else False for trace in fig.data]},
                           {"title": "Achat Marché Nature investissemnt - 2022"}]),
                dict(label="2023",
                     method="update",
                     args=[{"visible": [True if '2023' in trace.name else False for trace in fig.data]},
                           {"title": "Achat Marché Nature investissemnt - 2023"}]),
                dict(label=str(current_year),
                            method="update",
                         args=[{"visible": [True if f'{current_year}' in trace.name else False for trace in fig.data]},
                           {"title": f"Investment Data - {current_year}"}])
                ]
            )
            ],
                    title="Achat Marché investi Nature by Year",
                    xaxis_title="Nature",
                    yaxis_title="Value",
                    barmode='stack'
                )

           # Display the interactive stacked bar chart in Streamlit
        st.plotly_chart(fig)
                # Charger les données des feuilles de calcul
        achat_cdc_2021 = pd.read_excel(excel_data, sheet_name='Achat CDC par Budget2021')
        achat_cdc_2022 = pd.read_excel(excel_data, sheet_name='Achat CDC par Budget 2022')
        achat_cdc_2023 = pd.read_excel(excel_data, sheet_name='Achat CDC par Budget 2023')

         # Ajouter les colonnes 'Year'
        achat_cdc_2021['Year'] = 2021
        achat_cdc_2022['Year'] = 2022
        achat_cdc_2023['Year'] = 2023

        # Combiner les données en un seul DataFrame
        achat_cdc_all_years = pd.concat([achat_cdc_2021, achat_cdc_2022, achat_cdc_2023])

          # Créer la figure du graphique en barres empilées
        fig = go.Figure()

      # Ajouter des traces pour chaque année
        for year in [2021, 2022, 2023]:
                df_year = achat_cdc_all_years[achat_cdc_all_years['Year'] == year]
                fig.add_trace(go.Bar(
                x=df_year['Budget'],
                y=df_year['Nombre de CDC '],
                name=f'Nombre de CDC - {year}',
                marker_color='rgba(55, 83, 109, 0.7)'
              ))
                fig.add_trace(go.Bar(
                x=df_year['Budget'],
                y=df_year['Montant en Dhs'],
                name=f'Montant en Dhs - {year}',
                marker_color='rgba(255, 144, 14, 0.7)'
              ))

               # Mettre à jour la mise en page pour l'apparence des barres empilées
                fig.update_layout(
                barmode='stack',
                title="Achat CDC par Budget par Année",
                xaxis_title="Budget",
                yaxis_title="Valeurs (Nombre de CDC / Montant en Dhs)",
               legend_title="Année",
               )

               # Afficher le graphique interactif dans Streamlit
        st.plotly_chart(fig)
               
                # Load data from the sheets
       
        synthese_2023 = pd.read_excel(excel_data, sheet_name='Synthese par type 2023')

                # Add 'Year' columns
        
        synthese_2023['Year'] = 2023

                # Combine the data into one dataframe
        synthese_all_years = pd.concat([ synthese_2023])

                # Create the donut chart figure
        fig = go.Figure()

               # Add donut chart traces for each year
        for year in [ 2023]:
                df_year = synthese_all_years[synthese_all_years['Year'] == year]
                fig.add_trace(go.Pie(
                labels=df_year['Type'],
                values=df_year['Marché'],  # Replace with the column you want to visualize
                name=f'{year}',
                hole=0.4,  # Creates the donut shape
                hoverinfo='label+percent+name',  # Shows the label, percentage, and year on hover
                title=f"Année {year}",
               
            ))

                # Update layout to arrange the donut charts
        fig.update_layout(
               title_text="Synthese par Type Data -  by Year",
            annotations=[
            
                dict(text='', x=0.83, y=0.5, font_size=12, showarrow=False)
            ],
               showlegend=True
           )

         # Show the interactive donut chart in Streamlit
        st.plotly_chart(fig)

        # Load data from the sheets
        recettes_fonctionnement = pd.read_excel(excel_data, sheet_name='Recettes_Fonctionnement_2023')
        recettes_investissement = pd.read_excel(excel_data, sheet_name='Recettes_Invistissement_2023')
        recettes_propres = pd.read_excel(excel_data, sheet_name='RecettesPropres2023')

        # Add 'Category' columns
        recettes_fonctionnement['Category'] = 'Fonctionnement'
        recettes_investissement['Category'] = 'Investissement'
        recettes_propres['Category'] = 'Propres'

        # Combine the data into one dataframe
        recettes_all = pd.concat([recettes_fonctionnement, recettes_investissement, recettes_propres])
        
       # Create a filter for categories
        selected_categories = st.multiselect('Select Categories',
         options=recettes_all['Category'].unique(),
         default=recettes_all['Category'].unique()  # Default to all categories
        )

        # Filter the data based on selected categories
        filtered_data = recettes_all[recettes_all['Category'].isin(selected_categories)]

        # Create the bar chart figure
        fig = go.Figure()

        # Add traces for each category
        for category in recettes_all['Category'].unique():
          df_category = recettes_all[recettes_all['Category'] == category]
        if category == 'Propres':
          fig.add_trace(go.Bar(
            x=df_category['Désignation'] if 'Désignation' in df_category.columns else df_category['Subvention'],
            y=df_category['Montant en Dhs'],
            name=f'{category} - Montant en Dhs'
            ))
        else:
          fig.add_trace(go.Bar(
            x=df_category['Subvention'],
            y=df_category['Montant en dhs'],
            name=f'{category} - Montant en dhs'
           ))

         # Update the layout with title and labels
        fig.update_layout(
        title='Revenues by Category for 2023',
        xaxis_title='Category',
        yaxis_title='Montant en Dhs',
         barmode='stack'  # Use 'group' if you prefer separate bars for each category
        )

             # Display the interactive bar chart in Streamlit
        st.plotly_chart(fig)

    except Exception as e:
        st.error(f"An error occurred: {e}")


    
elif data_type == "Ressources humaines":
    try:
        # Lecture des données RH
        excel_data = pd.ExcelFile(uploaded_file)
            # Charger les données RH
        repart_grade = pd.read_excel(excel_data, sheet_name='Répartition par Grade')
        repart_genre = pd.read_excel(excel_data, sheet_name='Répartition par genre')
        repart_age = pd.read_excel(excel_data, sheet_name='Répartition par Age')
        repart_departement = pd.read_excel(excel_data, sheet_name='Répartition par Département') 
        repart_division = pd.read_excel(excel_data, sheet_name='Repartition par Division') 
        df_disponibilite = pd.read_excel(excel_data, sheet_name='Mise en disponibilité')  
        df_mise_disposition = pd.read_excel(excel_data, sheet_name='Mise à la disposistion') 
        df_detachements = pd.read_excel(excel_data, sheet_name='Détéchements')  
        df_promotions = pd.read_excel(excel_data, sheet_name='Promotions') 
        df_stages = pd.read_excel(excel_data ,sheet_name='Stages')
        df_mutation = pd.read_excel(excel_data ,sheet_name='Mutation permutation')
        df_retraite_grade = pd.read_excel(excel_data , sheet_name='Retraite par grade')
        df_recrutment = pd.read_excel(excel_data , sheet_name='Recrutment depuis 2015')
        df_diplome = pd.read_excel(excel_data ,sheet_name='Repartition par diplome')
            # Calculer les indicateurs
        total_effectif = repart_grade['Nombre'].sum()
        total_hommes = repart_genre[repart_genre['Genre'] == 'Homme']['Nombre'].sum()
        total_femmes = repart_genre[repart_genre['Genre'] == 'Femme']['Nombre'].sum()
        pourcentage_hommes = (total_hommes / total_effectif) * 100
        pourcentage_femmes = (total_femmes / total_effectif) * 100
            # Ajouter une section avec le logo comme icône avant le titre
        
        st.markdown("<h1 style='text-align: left;'>Indicateurs Clés</h1>", unsafe_allow_html=True)

            # Afficher les indicateurs dans des colonnes
        col1, col2, col3 = st.columns(3)

        with col1:
                st.info("Total Effectif", icon="📊")
                st.metric(label="Total", value=total_effectif)

        with col2:
                st.info("Total Hommes", icon="👨")
                st.metric(label="Nombre", value=total_hommes, delta=f"{pourcentage_hommes:.2f}%")

        with col3:
                st.info("Total Femmes", icon="👩")
                st.metric(label="Nombre", value=total_femmes, delta=f"{pourcentage_femmes:.2f}%")
             
            # Visualiser "Répartition par Grade"
        selected_categories = st.multiselect(
                "Sélectionnez les catégories à afficher",
                options=repart_grade['Catégorie'].unique(),
                default=list(repart_grade['Catégorie'].unique())
            )
        filtered_data_grade = repart_grade[repart_grade['Catégorie'].isin(selected_categories)]
        fig_grade = px.bar(
                filtered_data_grade,
                x='Catégorie',
                y='Nombre',
                color='Pourcentage %',
                title="Répartition par Grade"
            )
        st.plotly_chart(fig_grade)

            # Visualiser "Répartition par Genre"
        selected_genres = st.multiselect(
                "Sélectionnez les genres à afficher",
                options=repart_genre['Genre'].unique(),
                default=list(repart_genre['Genre'].unique())
            )
        filtered_data_genre = repart_genre[repart_genre['Genre'].isin(selected_genres)]
        fig_genre = px.pie(
                filtered_data_genre,
                names='Genre',
                values='Nombre',
                title="Répartition par Genre",
                hover_data=['Pourcentage %'],
                labels={'Pourcentage %': '% de Genre'}
            )
        st.plotly_chart(fig_genre)

            # Visualiser "Répartition par Tranche d'Âge"
        selected_age_ranges = st.multiselect(
                "Sélectionnez les tranches d'âge à afficher",
                options=repart_age["tranche d'âge"].unique(),
                default=list(repart_age["tranche d'âge"].unique())
            )
        filtered_data_age = repart_age[repart_age["tranche d'âge"].isin(selected_age_ranges)]
        fig_age = px.bar(
                filtered_data_age,
                x="tranche d'âge",
                y='Effectif',
                color='%',
                title="Répartition par Tranche d'Âge"
            )
        st.plotly_chart(fig_age)

            # Visualiser "Répartition par Département"
        st.subheader("Répartition par Département")
        fig_departement = px.pie(
            repart_departement, 
            names='Entité', 
            values='nombre du personnel', 
            title="Répartition par Département",
            labels={'nombre du personnel': 'Nombre du Personnel'}
        )
        st.plotly_chart(fig_departement)
            # Visualiser "Répartition par Division"
        st.subheader("Répartition par Division")
        fig_division = px.bar(
            repart_division, 
            x='Divisions', 
            y='Effectifs', 
            color='%',
            title="Répartition par Division",
            labels={'Effectifs': 'Effectifs', 'Divisions': 'Division'}
        )
        st.plotly_chart(fig_division)
        # Visualiser "Mise en Disponibilité"
        st.subheader("Mise en Disponibilité")
        fig_disponibilite = px.scatter(
            df_disponibilite, 
            x='Grade', 
            y='Unité', 
            color='Motif',
            title="Mise en Disponibilité",
            labels={'Grade': 'Grade', 'Unité': 'Unité'}
        )
        st.plotly_chart(fig_disponibilite)
        # Visualiser "Mise à la Disposition"
        st.subheader("Mise à la Disposition")
        fig_mise_disposition = px.scatter(
            df_mise_disposition, 
            x='Grade ', 
            y='Unité', 
            color="Administration d'accueil",
            title="Mise à la Disposition",
            labels={'Grade ': 'Grade', 'Unité': 'Unité'}
        )
        st.plotly_chart(fig_mise_disposition)
        # Visualiser "Détachements"
        st.subheader("Détachements")
        fig_detachements = px.scatter(
            df_detachements, 
            x='Grade', 
            y='Unité', 
            color="Administration d'accueil",
            title="Détachements",
            labels={'Grade': 'Grade', 'Unité': 'Unité'}
        )
        st.plotly_chart(fig_detachements)
        # Visualiser "Promotions"
        st.subheader("Promotions")
        df_promotions_long = df_promotions.melt(
            id_vars=['Cadre'], 
            value_vars=['Promotion de grade', 'Avancement d’échelon', 'Notation', 'Titularisation'],
            var_name='Type de Promotion', 
            value_name='Nombre'
        )
        fig_promotions = px.bar(
            df_promotions_long, 
            x='Cadre', 
            y='Nombre', 
            color='Type de Promotion',
            title="Distribution des Promotions par Cadre"
        )
        st.plotly_chart(fig_promotions)
        # Visualisation pour "Répartition des Stages"
        df_stages_long = df_stages.melt(var_name='Stage', value_name='Nombre')
        df_stages_long = df_stages_long[df_stages_long['Stage'].str.contains('Unnamed') == False]
        fig_stages = px.bar(df_stages_long, x='Stage', y='Nombre', 
                    title="Répartition des Stages",
                    labels={'Stage': 'Stage', 'Nombre': 'Nombre'})
        st.plotly_chart(fig_stages)

         # Visualisation pour "Mutations par Grade"
        fig_mutation = px.bar(df_mutation, x='Grade', title="Mutations par Grade",
                      labels={'Grade': 'Grade', 'Unnamed: 0': 'Nombre'})
        st.plotly_chart(fig_mutation)
        # Visualisation pour "Départs à la Retraite par Catégorie"
        fig_retraite_grade = px.bar(df_retraite_grade, x='Catégorie', y=' Départs à la retraite',
                           title="Départs à la Retraite par Catégorie",
                           labels={'Catégorie': 'Catégorie', ' Départs à la retraite': 'Départs à la Retraite'})
        st.plotly_chart(fig_retraite_grade)
        # Visualisation pour "Recrutement par Année"
        fig_recrutment = px.line(df_recrutment, x='ANNEE', y='NOMBRE', color='CADRE',
                            title="Recrutement par Année",
                            labels={'ANNEE': 'Année', 'NOMBRE': 'Nombre de Recrutements', 'CADRE': 'Cadre'})
        st.plotly_chart(fig_recrutment)
        # Visualisation pour "Répartition des Diplômes"
        df_diplome_long = df_diplome.melt(var_name='Diplôme', value_name='Nombre')
        df_diplome_long = df_diplome_long.dropna()
        fig_diplome = px.pie(df_diplome_long, names='Diplôme', values='Nombre',
                        title="Répartition des Diplômes",
                        labels={'Nombre': 'Nombre', 'Diplôme': 'Diplôme'})
        st.plotly_chart(fig_diplome)
    except Exception as e:
            st.error(f"Une erreur est survenue lors du traitement des données RH : {e}")

             
# Footer
st.sidebar.markdown('''
---
Created with ❤️ by Louhou Hajar
''')
