import pandas as pd
from pptx import Presentation
from utils import add_competitor_row

data_for_df = [
    {
        "societa": "Beta Srl",
        "sede": "Via Verdi, 123 - Milano",
        "descrizione": "Opera nel settore food & beverage, con una lunga tradizione e prodotti di alta qualità riconosciuti sul mercato nazionale ed estero.",
        "vdp": "12.000",
        "ebitda": "3.000",
        "ebitda_percent": "25%",
        "n_dip": "60",
        "pfn": "900",
        "fcf": "250"
    },
    {
        "societa": "Gamma Spa",
        "sede": "Corso Italia, 45 - Roma",
        "descrizione": "Leader nella logistica urbana e trasporti sostenibili, con focus sull'innovazione tecnologica.",
        "vdp": "8.000",
        "ebitda": "1.600",
        "ebitda_percent": "20%",
        "n_dip": "40",
        "pfn": "1.200",
        "fcf": "180"
    },
    {
        "societa": "Delta Group International",
        "sede": "Piazza Affari, 1 - Torino",
        "descrizione": "Specializzati in consulenza finanziaria strategica e operazioni di M&A cross-border.",
        "vdp": "15.000",
        "ebitda": "4.500",
        "ebitda_percent": "30%",
        "n_dip": "75",
        "pfn": "500",
        "fcf": "400"
    },
    {
        "societa": "Epsilon Solutions",
        "sede": "Via Mazzini, 7 - Firenze",
        "descrizione": "Fornitore di software gestionali per PMI, con soluzioni cloud innovative e scalabili.",
        "vdp": "7.500",
        "ebitda": "1.800",
        "ebitda_percent": "24%",
        "n_dip": "35",
        "pfn": "600",
        "fcf": "220"
    },
    {
        "societa": "Zeta Corp",
        "sede": "Largo Garibaldi, 22 - Napoli",
        "descrizione": "Produzione e distribuzione di componenti elettronici per il settore automotive e industriale.",
        "vdp": "20.000",
        "ebitda": "5.000",
        "ebitda_percent": "25%",
        "n_dip": "100",
        "pfn": "2.000",
        "fcf": "500"
    },
    {
        "societa": "Eta Systems",
        "sede": "Viale Europa, 90 - Bologna",
        "descrizione": "Integrazione di sistemi di automazione industriale e robotica avanzata per manifattura.",
        "vdp": "9.000",
        "ebitda": "2.250",
        "ebitda_percent": "25%",
        "n_dip": "45",
        "pfn": "750",
        "fcf": "300"
    },
    {
        "societa": "Theta Industries",
        "sede": "Via Roma, 1 - Genova",
        "descrizione": "Costruzioni navali e offshore, specializzata in imbarcazioni da lavoro e piattaforme.",
        "vdp": "30.000",
        "ebitda": "6.000",
        "ebitda_percent": "20%",
        "n_dip": "150",
        "pfn": "5.000",
        "fcf": "800"
    },
    {
        "societa": "Iota Renewables",
        "sede": "Strada del Sole, 55 - Bari",
        "descrizione": "Sviluppo e gestione di impianti per la produzione di energia da fonti rinnovabili.",
        "vdp": "18.000",
        "ebitda": "7.200",
        "ebitda_percent": "40%",
        "n_dip": "80",
        "pfn": "3.000",
        "fcf": "1.000"
    },
    {
        "societa": "Kappa Logistics",
        "sede": "Piazza Duomo, 3 - Palermo",
        "descrizione": "Servizi di trasporto merci intermodale e gestione magazzini conto terzi su scala nazionale.",
        "vdp": "11.000",
        "ebitda": "1.980",
        "ebitda_percent": "18%",
        "n_dip": "55",
        "pfn": "1.500",
        "fcf": "150"
    },
    {
        "societa": "Lambda Biotech",
        "sede": "Via della Scienza, 10 - Trieste",
        "descrizione": "Ricerca e sviluppo nel campo delle biotecnologie applicate alla farmaceutica e agrifood.",
        "vdp": "25.000",
        "ebitda": "6.250",
        "ebitda_percent": "25%",
        "n_dip": "90",
        "pfn": "2.500",
        "fcf": "700"
    },
    {
        "societa": "Mu Digital",
        "sede": "Corso Vittorio Emanuele, 70 - Cagliari",
        "descrizione": "Agenzia di marketing digitale specializzata in SEO, SEM e social media strategy per aziende.",
        "vdp": "5.000",
        "ebitda": "1.500",
        "ebitda_percent": "30%",
        "n_dip": "25",
        "pfn": "300",
        "fcf": "100"
    },
    {
        "societa": "Nu Innovations",
        "sede": "Via dell'Innovazione, 42 - Trento",
        "descrizione": "Startup focalizzata su soluzioni IoT per smart cities e monitoraggio ambientale.",
        "vdp": "6.000",
        "ebitda": "1.200",
        "ebitda_percent": "20%",
        "n_dip": "30",
        "pfn": "400",
        "fcf": "80"
    },
    {
        "societa": "Xi Holdings",
        "sede": "Via Finanza, 8 - Milano",
        "descrizione": "Società di investimento con portafoglio diversificato in vari settori industriali e tecnologici.",
        "vdp": "50.000",
        "ebitda": "12.500",
        "ebitda_percent": "25%",
        "n_dip": "200",
        "pfn": "8.000",
        "fcf": "2.000"
    },
    {
        "societa": "Omicron Services",
        "sede": "Largo dei Servizi, 15 - Ancona",
        "descrizione": "Servizi di consulenza aziendale per l'ottimizzazione dei processi e la trasformazione digitale.",
        "vdp": "13.000",
        "ebitda": "3.250",
        "ebitda_percent": "25%",
        "n_dip": "65",
        "pfn": "1.100",
        "fcf": "350"
    },
    {
        "societa": "Pi Manufacturing",
        "sede": "Zona Industriale Est, Blocco 5 - Perugia",
        "descrizione": "Produzione di macchinari industriali su misura per il settore metalmeccanico e plastico.",
        "vdp": "16.000",
        "ebitda": "3.200",
        "ebitda_percent": "20%",
        "n_dip": "70",
        "pfn": "2.200",
        "fcf": "450"
    }
]

competitors_df = pd.DataFrame(data_for_df)

prs = Presentation("Competitors.pptx")

model_slide_layout = prs.slides[1].slide_layout
current_slide = prs.slides[1]

offset_per_riga_emu = 763200
items_per_slide = 6
rows_written_on_current_slide = 0

if not competitors_df.empty:
    for index, competitor_data_dict in enumerate(competitors_df.to_dict(orient='records')):
        if index > 0 and index % items_per_slide == 0:
            current_slide = prs.slides.add_slide(model_slide_layout) #TODO: crare nuova funuzuione per slide con gli elementi
            rows_written_on_current_slide = 0

        current_y_offset = rows_written_on_current_slide * offset_per_riga_emu
        
        add_competitor_row(current_slide, competitor_data_dict, y_offset=current_y_offset)
        
        rows_written_on_current_slide += 1
else:
    print("Il DataFrame è vuoto. Nessuna riga da aggiungere.")

prs.save("Competitors_modificato_pandas.pptx")
