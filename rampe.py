import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import os
import pandas as pd 
from datetime import datetime
from reportlab.pdfgen.canvas import Canvas # Importa esplicitamente Canvas

# Importazione di reportlab, necessaria per la generazione PDF.
try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.pdfgen.canvas import Canvas 

except ImportError:
    print("AVVISO: La libreria 'reportlab' non è installata. La generazione PDF sarà disabilitata.")


# --- Costanti e Utilità ---
NS_URI = "http://www.orienteering.org/datastandard/3.0"
def q(tag):
    return f"{{{NS_URI}}}{tag}"
OUTPUT_CONFIG_FILE = "configurazione.cfg"
OUTPUT_FILE_INDIVIDUALE = "classifica_rampe_per_categoria.xlsx"
OUTPUT_FILE_STAFFETTE = "classifica_staffette_rampe.xlsx"
OUTPUT_FILE_PDF_INDIVIDUALE = "classifica_rampe_individuali.pdf"
OUTPUT_FILE_PDF_STAFFETTE = "classifica_rampe_staffette.pdf"

# --- FUNZIONI DI UTILITY (omesse per brevità, sono invariate) ---

def calcola_codice_staffetta(pettorale_str):
    try:
        pettorale_int = int(pettorale_str)
        if pettorale_int > 999:
            return "IND"
        elif pettorale_int <= 999:
            formatted = f"{pettorale_int:03d}"
            return formatted[:2]
        else:
            return "IND" 
    except ValueError:
        return "IND" 

def formatta_tempo_hhmmss(secondi):
    if pd.isna(secondi) or secondi is None or secondi == "PM":
        return "PM"
    try:
        secondi = int(secondi)
        ore = secondi // 3600
        minuti = (secondi % 3600) // 60
        sec = secondi % 60
        return f"{ore:02d}:{minuti:02d}:{sec:02d}"
    except:
        return str(secondi)
    
def analyze_relay_teams(df_individuale):
    if df_individuale.empty:
        return pd.DataFrame(), []

    df_valid = df_individuale[
        (df_individuale['Somma'] != "PM") & 
        (df_individuale['Codice_Staffetta'] != "IND")
    ].copy()
    
    if df_valid.empty:
        return pd.DataFrame(), []
        
    df_valid['Somma_float'] = df_valid['Somma']
    
    # Identifica tutte le colonne di rampa (Δ)
    delta_cols = [col for col in df_valid.columns if col.startswith('Δ')]
    
    grouped = df_valid.groupby('Codice_Staffetta')

    team_data = []
    avvisi = []

    for codice_staffetta, group in grouped:
        if len(group) < 2:
            avvisi.append(f"• Staffetta {codice_staffetta} ha solo {len(group)} frazionista(i) e viene ignorata.")
            continue
        
        group = group.sort_values(by='Pettorale').reset_index(drop=True)
        
        p1 = group.iloc[0]
        p2 = group.iloc[1]

        # IL CALCOLO È CORRETTO: somma le colonne 'Somma_float' che contengono il totale di tutte le rampe individuali
        total_ramp_time = p1['Somma_float'] + p2['Somma_float']
        
        team_name = f"{p1['Nome']} / {p2['Nome']}"

        # Inizializza la riga della squadra
        team_row = {
            "Codice Staffetta": codice_staffetta,
            "Coppia Staffetta": team_name,
            "Tempo_Totale_Coppia": total_ramp_time
        }
        
        # Aggiungi i dettagli del Frazionista 1
        team_row["Pettorale_1"] = p1['Pettorale']
        team_row["Nome_1"] = p1['Nome']
        team_row["Totale_Rampe_1"] = p1['Somma_float'] # Totale di tutte le rampe per P1
        # Aggiungi TUTTE le colonne Δ per il Frazionista 1
        for col in delta_cols:
             team_row[f"{col}_P1"] = p1[col] 
        
        # Aggiungi i dettagli del Frazionista 2
        team_row["Pettorale_2"] = p2['Pettorale']
        team_row["Nome_2"] = p2['Nome']
        team_row["Totale_Rampe_2"] = p2['Somma_float'] # Totale di tutte le rampe per P2
        # Aggiungi TUTTE le colonne Δ per il Frazionista 2
        for col in delta_cols:
             team_row[f"{col}_P2"] = p2[col]

        team_data.append(team_row)
        
        if len(group) > 2:
             avvisi.append(f"• Staffetta {codice_staffetta} ha {len(group)} frazionisti. Considerati solo i primi 2 ({p1['Pettorale']} e {p2['Pettorale']}).")


    df_teams = pd.DataFrame(team_data)

    if not df_teams.empty:
        df_teams = df_teams.sort_values(by='Tempo_Totale_Coppia').reset_index(drop=True)
        df_teams['Posizione'] = df_teams.index + 1
        
        # Riordina le colonne in modo flessibile per l'Excel
        base_cols = ['Posizione', 'Codice Staffetta', 'Coppia Staffetta', 'Tempo_Totale_Coppia']
        
        # Colonne Frazionista 1: Pettorale, Nome, Totale, poi i dettagli di tutte le Δ
        p1_cols = [f"Pettorale_1", f"Nome_1", f"Totale_Rampe_1"]
        p1_cols.extend([f"{col}_P1" for col in delta_cols])
        
        # Colonne Frazionista 2: Pettorale, Nome, Totale, poi i dettagli di tutte le Δ
        p2_cols = [f"Pettorale_2", f"Nome_2", f"Totale_Rampe_2"]
        p2_cols.extend([f"{col}_P2" for col in delta_cols])
        
        all_cols = base_cols + p1_cols + p2_cols
        
        # Seleziona solo le colonne che esistono nel DataFrame (garanzia)
        df_teams = df_teams[[c for c in all_cols if c in df_teams.columns]]
        

    return df_teams, avvisi

# --------------------------------------------------------------------------------------------------
# --- CLASSE CANVAS CORRETTA (FIX DUPLICAZIONE) ---
# --------------------------------------------------------------------------------------------------

class PageNumberCanvas(Canvas):
    """
    Canvas personalizzata per inserire il numero totale di pagine (N) in un secondo
    momento, evitando la scrittura doppia del contenuto della pagina.
    """
    def __init__(self, *args, **kwargs):
        # Estrai il print_time e le impostazioni del SimpleDocTemplate
        self.print_time = kwargs.pop('print_time', None)
        self.doc_settings = kwargs.pop('doc_settings', {})
        Canvas.__init__(self, *args, **kwargs)
        self.pages = [] # Usato per salvare lo stato di ogni pagina
        
        # Salva gli argomenti passati a __init__ (necessari per il reset)
        self._args = args
        self._kwargs = kwargs

    def showPage(self):
        """Salva lo stato della pagina, ma non la scrive nel PDF per evitare duplicazioni."""
        
        # 1. Salva lo stato completo del canvas (incluso il contenuto disegnato dai flowables)
        # Usiamo copy.copy per evitare riferimenti incrociati
        import copy
        self.pages.append(copy.copy(self.__dict__))
        
        # 2. Resetta lo stato del canvas per la pagina successiva
        # Si usa l'inizializzazione con gli argomenti originali, ma si evita di scrivere la pagina
        self._startPage() 

    def save(self):
        """Itera attraverso gli stati delle pagine salvate, disegna il footer finale e scrive la pagina."""
        page_count = len(self.pages)
        for page_index, page_data in enumerate(self.pages):
            
            # 1. Ripristina lo stato salvato (il contenuto)
            self.__dict__.update(page_data) 
            
            # 2. Imposta il numero di pagina per il piè di pagina (1-based index)
            self._pageNumber = page_index + 1 
            self.draw_footer(page_count) # Disegna il footer finale
            
            # 3. Scrivi la pagina *una sola volta* con il footer incluso
            Canvas.showPage(self) 
            
        Canvas.save(self)
        
    def draw_footer(self, page_count):
        """Logica di disegno del piè di pagina."""
        canvas = self
        page_num = canvas.getPageNumber()
        page_width, _ = landscape(A4)
        
        # Recupera i margini dai settings
        leftMargin = self.doc_settings.get('leftMargin', 1*cm)
        rightMargin = self.doc_settings.get('rightMargin', 1*cm)
        y_pos = 1 * cm
        
        canvas.setFont('Helvetica', 9)
        
        # 1. Ora di Stampa (Sinistra)
        text_stampa = f"Ora di stampa: {self.print_time}"
        canvas.drawString(leftMargin, y_pos, text_stampa)
        
        # 2. Numero di pagina (P / N) (Destra)
        footer_text = f"Pagina {page_num} / {page_count}"
        canvas.drawRightString(page_width - rightMargin, y_pos, footer_text)


# --------------------------------------------------------------------------------------------------
# --- FUNZIONI DI GENERAZIONE PDF (AGGIORNATE PER USARE PageNumberCanvas) ---
# Le funzioni PDF rimangono INVARIATE come richiesto
# --------------------------------------------------------------------------------------------------

def genera_pdf_individuale(output_sheets, print_time, filename=OUTPUT_FILE_PDF_INDIVIDUALE):
# ... (invariata)
    if not output_sheets:
        return
    
    doc_settings = {'leftMargin': 1*cm, 'rightMargin': 1*cm, 'topMargin': 1.5*cm, 'bottomMargin': 1.5*cm}
        
    doc = SimpleDocTemplate(filename, pagesize=landscape(A4), **doc_settings)
    styles = getSampleStyleSheet()
    story = []

    categorie_list = list(output_sheets.keys())
    
    for idx, cat in enumerate(categorie_list):
        df = output_sheets[cat]
        if df.empty:
            continue
            
        story.append(Paragraph(f"<b>Classifica individuale cronoscalata categoria: {cat}</b>", styles['Title']))
        story.append(Spacer(1, 0.4*cm))

        df_pdf = df.copy()
        delta_cols = [col for col in df_pdf.columns if col.startswith('Δ')]
        df_pdf['Somma_HHMMSS'] = df_pdf['Somma'].apply(formatta_tempo_hhmmss)
        cols_for_pdf = ['Pettorale', 'Nome'] + delta_cols + ['Somma_HHMMSS']
        df_pdf = df_pdf[cols_for_pdf].reset_index(drop=True)
        df_pdf.insert(0, 'Posizione', df_pdf.index + 1)
        
        new_cols = ['Pos.', 'Pett.', 'Nome Atleta'] + [col.replace('Δ', 'Diff. ') for col in delta_cols] + ['Totale']
        df_pdf.columns = new_cols
        
        data = [df_pdf.columns.tolist()] + df_pdf.values.tolist()
        
        num_cols = len(df_pdf.columns)
        col_widths = [1.2*cm, 1.2*cm, 6*cm]
        remaining_width = landscape(A4)[0] - sum(col_widths) - (2*cm)
        dynamic_width = remaining_width / (num_cols - 3)
        col_widths += [max(1.5*cm, dynamic_width)] * (num_cols - 3)
        col_widths = [min(w, 4*cm) for w in col_widths] 

        table = Table(data, colWidths=col_widths) 
        
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (2, 0), (2, -1), 'LEFT'), 
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 8) 
        ])
        
        table.setStyle(style)
        story.append(table)
        
        if idx < len(categorie_list) - 1:
            story.append(Spacer(1, 0.5*cm))
            story.append(PageBreak()) 

    try:
        if story:
            doc.build(story, 
                      canvasmaker=lambda *args, **kwargs: PageNumberCanvas(*args, 
                                                                           print_time=print_time, 
                                                                           doc_settings=doc_settings,
                                                                           **kwargs))
            return True
        return False
    except NameError:
        messagebox.showwarning("Avviso PDF", "Impossibile generare il PDF individuale. La libreria 'reportlab' non è installata.")
        return False
    except Exception as e:
        messagebox.showerror("Errore PDF", f"Errore durante la creazione del PDF individuale: {e}")
        return False


def genera_pdf_staffette(relay_sheets, print_time, filename=OUTPUT_FILE_PDF_STAFFETTE):
# ... (invariata)
    if not relay_sheets:
        return

    doc_settings = {'leftMargin': 1*cm, 'rightMargin': 1*cm, 'topMargin': 1.5*cm, 'bottomMargin': 1.5*cm}

    doc = SimpleDocTemplate(filename, pagesize=landscape(A4), **doc_settings)
    styles = getSampleStyleSheet()
    story = []

    categorie_list = list(relay_sheets.keys())
    
    for idx, cat in enumerate(categorie_list):
        df = relay_sheets[cat]
        if df.empty:
            continue
            
        story.append(Paragraph(f"<b>Classifica di squadra sulle cronoscalate - Categoria: {cat}</b>", styles['Title']))
        story.append(Spacer(1, 0.4*cm))

        df_pdf = df.copy()
        
        df_pdf['Tempo_Totale_Coppia_HHMMSS'] = df_pdf['Tempo_Totale_Coppia'].apply(formatta_tempo_hhmmss)

        # Il PDF usa ancora solo Nome_1 e Nome_2 come da richiesta di non modificarlo
        df_pdf = df_pdf[['Posizione', 'Nome_1', 'Nome_2', 'Tempo_Totale_Coppia_HHMMSS']]
        df_pdf.columns = ['Pos.', 'Concorrente 1', 'Concorrente 2', 'Tempo Totale']
        
        data = [df_pdf.columns.tolist()] + df_pdf.values.tolist()
        
        table = Table(data, colWidths=[1.5*cm, 8*cm, 8*cm, 4*cm])
        
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (1, 0), (2, -1), 'LEFT'), 
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.lightblue),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ])
        
        table.setStyle(style)
        story.append(table)
        
        if idx < len(categorie_list) - 1:
            story.append(Spacer(1, 0.5*cm))
            story.append(PageBreak()) 

    try:
        if story:
            doc.build(story,
                      canvasmaker=lambda *args, **kwargs: PageNumberCanvas(*args, 
                                                                           print_time=print_time, 
                                                                           doc_settings=doc_settings,
                                                                           **kwargs))
            return True
        return False
    except NameError:
        messagebox.showwarning("Avviso PDF", "Impossibile generare il PDF staffette. La libreria 'reportlab' non è installata.")
        return False
    except Exception as e:
        messagebox.showerror("Errore PDF", f"Errore durante la creazione del PDF staffette: {e}")
        return False


# --- CLASSE APPLICAZIONE GUI (OMESSA PER BREVITÀ, È INVARIATA) ---
class App(tk.Tk):
# ... (Il resto della classe App è invariato) ...
    def __init__(self):
        super().__init__()
        self.title("SkyRun generatore classifiche Rampe - By Fox 2025")
        self.geometry("650x700")
        self.minsize(550, 400)
        
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('Accent.TButton', foreground='blue', font=('TkDefaultFont', 10, 'bold')) 
        
        self.file_path = None
        self.categorie = []
        self.widgets_rampe = {}
        self.configurazione_caricata = {}

        self._carica_configurazione_iniziale()
        self.create_widgets()
        
        if self.configurazione_caricata:
            print(f"Configurazione caricata da '{OUTPUT_CONFIG_FILE}'.")
            
    def _carica_configurazione_iniziale(self):
        if os.path.exists(OUTPUT_CONFIG_FILE):
            try:
                with open(OUTPUT_CONFIG_FILE, 'r') as f:
                    self.configurazione_caricata = json.load(f)
                self.categorie = list(self.configurazione_caricata.keys())
            except Exception:
                messagebox.showwarning("Attenzione", f"Il file '{OUTPUT_CONFIG_FILE}' è corrotto o non leggibile. Verrà ignorato.")
                self.configurazione_caricata = {}
                self.categorie = []


    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill='both', expand=True)
        
        file_control_frame = ttk.Frame(main_frame, padding="5")
        file_control_frame.pack(fill='x', pady=5)

        ttk.Button(file_control_frame, text="Seleziona XML", command=self.seleziona_xml).pack(side='left', padx=(0, 10))
        ttk.Button(file_control_frame, text="❌ Chiudi", command=self.destroy).pack(side='right', padx=(10, 0))

        self.file_label = ttk.Label(file_control_frame, text=f"File config: {OUTPUT_CONFIG_FILE} - XML: Nessuno")
        self.file_label.pack(side='left', fill='x', expand=True, padx=5)

        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=5)
        
        canvas = tk.Canvas(main_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, padding="5")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        self.rampe_container_frame = scrollable_frame

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        if self.categorie:
            self.ricrea_interfaccia_rampe(self.configurazione_caricata)
        else:
            ttk.Label(self.rampe_container_frame, text="Seleziona un file XML per iniziare la configurazione.", foreground='gray').pack(pady=20)


        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=5)
        
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill='x', pady=10)
        
        ttk.Button(action_frame, text="✔️ Salva Configurazione", 
                   command=lambda: self.salva_configurazione(silent=False), 
                   style='Accent.TButton').pack(fill='x', expand=True, padx=5, pady=2)

        ttk.Button(action_frame, text="🏆 Genera Classifica (Excel/PDF)", 
                   command=self.genera_classifica, 
                   style='Accent.TButton').pack(fill='x', expand=True, padx=5, pady=2)

    def ricrea_interfaccia_rampe(self, dati_rampe=None):
        for widget in self.rampe_container_frame.winfo_children():
            widget.destroy()
        self.widgets_rampe = {}
        
        if not self.categorie:
            ttk.Label(self.rampe_container_frame, text="Nessuna categoria trovata.", foreground='gray').pack(pady=20)
            return

        for cat in self.categorie:
            self.widgets_rampe[cat] = []
            rampe_prefill = dati_rampe.get(cat, []) if dati_rampe else []
            self.aggiungi_categoria_gui(cat, rampe_prefill)
            
        self.rampe_container_frame.update_idletasks()
        self.rampe_container_frame.master.config(scrollregion=self.rampe_container_frame.master.bbox("all"))

    def seleziona_xml(self):
        xml_file = filedialog.askopenfilename(
            title="Seleziona il file XML da cui estrarre le categorie",
            filetypes=[("File XML", "*.xml"), ("Tutti i file", "*.*")]
        )

        if xml_file:
            self.file_path = xml_file
            self.file_label.config(text=f"File config: {OUTPUT_CONFIG_FILE} - XML: {os.path.basename(xml_file)}")
            self.estrai_categorie(xml_file)
        else:
            messagebox.showinfo("Informazione", "Nessun file XML selezionato.")

    def estrai_categorie(self, xml_file):
        self.categorie = []

        try:
            tree = ET.parse(xml_file)
            root_xml = tree.getroot()
        except Exception as e:
            messagebox.showerror("Errore XML", f"Errore durante la lettura del file XML: {e}")
            return
        
        for classres in root_xml.findall(".//" + q("ClassResult")):
            cat = classres.findtext(q("Class") + "/" + q("Name"))
            if cat and cat not in self.categorie:
                self.categorie.append(cat)

        if not self.categorie:
            messagebox.showinfo("Informazione", "Nessuna categoria valida trovata nel file XML.")
            self.ricrea_interfaccia_rampe()
            return

        self.ricrea_interfaccia_rampe(self.configurazione_caricata)

    def aggiungi_categoria_gui(self, cat, rampe_prefill=[]):
        cat_frame = ttk.LabelFrame(self.rampe_container_frame, text=f" Categoria: {cat} ", padding="10")
        cat_frame.pack(fill='x', padx=5, pady=5)
        
        rampe_list_frame = ttk.Frame(cat_frame)
        rampe_list_frame.pack(fill='x', pady=5)

        ttk.Button(cat_frame, text="➕ Aggiungi Rampa", command=lambda: self.aggiungi_rampa_gui(cat, rampe_list_frame)).pack(pady=5)
        
        self.widgets_rampe[cat] = {'frame': rampe_list_frame, 'rampe': []}
        
        if rampe_prefill:
            for c1, c2 in rampe_prefill:
                self.aggiungi_rampa_gui(cat, rampe_list_frame, c1, c2)
        else:
            self.aggiungi_rampa_gui(cat, rampe_list_frame)

    def rimuovi_rampa_gui(self, cat, rampa_frame, widget_data):
        rampa_frame.destroy()
        if widget_data in self.widgets_rampe[cat]['rampe']:
            self.widgets_rampe[cat]['rampe'].remove(widget_data)
        
        self.rampe_container_frame.master.config(scrollregion=self.rampe_container_frame.master.bbox("all"))

    def aggiungi_rampa_gui(self, cat, parent_frame, c1_default="", c2_default=""):
        rampa_row = ttk.Frame(parent_frame, padding="5")
        rampa_row.pack(fill='x', padx=5, pady=2)
        
        ttk.Label(rampa_row, text="Da:").pack(side='left', padx=(0, 5))
        entry_start = ttk.Entry(rampa_row, width=8)
        entry_start.insert(0, str(c1_default))
        entry_start.pack(side='left', padx=5)

        ttk.Label(rampa_row, text="A:").pack(side='left', padx=(15, 5))
        entry_end = ttk.Entry(rampa_row, width=8)
        entry_end.insert(0, str(c2_default))
        entry_end.pack(side='left', padx=5)
        
        widget_data = {} 
        btn_elimina = ttk.Button(rampa_row, text="✖", width=2, command=lambda: self.rimuovi_rampa_gui(cat, rampa_row, widget_data))
        btn_elimina.pack(side='right')

        widget_data.update({
            'row_frame': rampa_row,
            'entry_start': entry_start,
            'entry_end': entry_end
        })
        
        self.widgets_rampe[cat]['rampe'].append(widget_data)
        self.rampe_container_frame.update_idletasks()
        self.rampe_container_frame.master.config(scrollregion=self.rampe_container_frame.master.bbox("all"))

    def salva_configurazione(self, silent=False):
        rampe_per_categoria = {}
        
        if not self.categorie:
            if not silent:
                messagebox.showwarning("Attenzione", "Devi prima selezionare un file XML o avere una configurazione caricata.")
            return None

        for cat in self.categorie:
            rampe_cat = []
            for rampa in self.widgets_rampe[cat]['rampe']:
                try:
                    c1_str = rampa['entry_start'].get().strip()
                    c2_str = rampa['entry_end'].get().strip()

                    if not c1_str and not c2_str:
                        continue
                        
                    c1 = int(c1_str)
                    c2 = int(c2_str)
                    
                    if c1 <= 0 or c2 <= 0:
                        raise ValueError("I codici devono essere numeri interi positivi.")

                    # CONTROLLO RILASSATO: Rimosso il controllo che c1 <= c2 per permettere qualsiasi sequenza di rampe.
                        
                    rampe_cat.append([c1, c2])
                    
                except ValueError as e:
                    messagebox.showerror("Errore di Input", f"Categoria {cat}: Inserisci codici numerici interi validi. ({e})")
                    return None
            
            rampe_per_categoria[cat] = rampe_cat

        try:
            with open(OUTPUT_CONFIG_FILE, 'w') as f:
                json.dump(rampe_per_categoria, f, indent=4)
            
            if not silent:
                messagebox.showinfo("Successo", f"Configurazione rampe completata e salvata in: {OUTPUT_CONFIG_FILE}")
            
            self.configurazione_caricata = rampe_per_categoria 
            return rampe_per_categoria
            
        except Exception as e:
            messagebox.showerror("Errore di Salvataggio", f"Errore durante il salvataggio del file di configurazione: {e}")
            return None


    def genera_classifica(self):
        """Esegue la logica di calcolo delle classifiche individuali, staffette ed i PDF."""
        
        rampe_per_categoria = self.salva_configurazione(silent=True)
        
        if rampe_per_categoria is None:
            return
            
        if not self.file_path:
            messagebox.showwarning("Attenzione", "Prima di generare la classifica, devi selezionare un file XML.")
            return
        
        if not any(rampe for rampe in rampe_per_categoria.values()):
            messagebox.showwarning("Attenzione", "Definisci le rampe per almeno una categoria prima di generare la classifica.")
            return

        try:
            tree = ET.parse(self.file_path)
            root_xml = tree.getroot()
        except Exception as e:
            messagebox.showerror("Errore XML", f"Impossibile leggere o interpretare il file XML: {e}")
            return

        output_sheets = {}   
        relay_sheets = {}    
        avvisi_staffette_generali = []
        
        # Ottieni l'ora attuale una sola volta per la stampa
        print_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S") 


        for classres in root_xml.findall(".//" + q("ClassResult")):
            categoria = classres.findtext(q("Class") + "/" + q("Name"))

            if categoria not in rampe_per_categoria or not rampe_per_categoria[categoria]:
                continue

            rampe = rampe_per_categoria[categoria]
            rows = []

            for person in classres.findall(q("PersonResult")):
                bib = person.findtext('.//' + q('BibNumber'))
                family = person.findtext('.//' + q('Family')) or ""
                given = person.findtext('.//' + q('Given')) or ""
                fullname = f"{family} {given}".strip()

                times = {}
                for split in person.findall('.//' + q('SplitTime')):
                    code = split.findtext(q('ControlCode'))
                    time = split.findtext(q('Time'))
                    if code and time:
                        try:
                            times[int(code)] = int(time)
                        except:
                            pass

                differenze = []
                dati_mancanti = False

                for c1_list, c2_list in rampe:
                    c1 = int(c1_list)
                    c2 = int(c2_list)
                    
                    if c1 in times and c2 in times:
                        differenze.append(times[c2] - times[c1])
                    else:
                        differenze.append(None)
                        dati_mancanti = True

                somma = sum(filter(lambda x: x is not None, differenze)) if not dati_mancanti else "PM"

                row = {"Pettorale": bib, "Nome": fullname}

                for idx, (c1, c2) in enumerate(rampe, start=1):
                    row[f"T{c1}"] = times.get(c1)
                    row[f"T{c2}"] = times.get(c2)
                    row[f"Δ{c2}-{c1}"] = differenze[idx-1] 

                row["Somma"] = somma
                row["Codice_Staffetta"] = calcola_codice_staffetta(bib)
                rows.append(row)

            df = pd.DataFrame(rows)
            
            if not df.empty:
                # 1. Classifica individuale
                df_indiv = df.copy()
                df_indiv["Somma_sort"] = df_indiv["Somma"].apply(lambda x: float("inf") if x == "PM" or pd.isna(x) else x)
                
                cols_order = ["Pettorale", "Codice_Staffetta", "Nome"] + [col for col in df_indiv.columns if col.startswith('Δ')] + ["Somma"]
                df_indiv = df_indiv.sort_values(by=["Somma_sort"]).drop(columns="Somma_sort")
                
                excel_cols = [col for col in df_indiv.columns if not col.startswith('T') or col.startswith('Pettorale') or col.startswith('Nome')]
                output_sheets[categoria] = df_indiv[[c for c in excel_cols if c in df_indiv.columns]]
                
                # 2. Classifica staffette
                df_relay, avvisi_cat = analyze_relay_teams(df)
                if not df_relay.empty:
                    relay_sheets[categoria] = df_relay
                    if avvisi_cat:
                         avvisi_staffette_generali.append(f"Avvisi Categoria {categoria}:")
                         avvisi_staffette_generali.extend([f"   {a}" for a in avvisi_cat])
                elif not df_indiv[df_indiv["Codice_Staffetta"] != "IND"].empty:
                    avvisi_staffette_generali.append(f"• Categoria {categoria}: Nessuna coppia staffetta valida trovata.")


        # 5. SCRITTURA EXCEL
        risultato_msg = ""
        try:
            # Funzione per sanificare i nomi dei fogli Excel
            def sanitize_sheet_name(name):
                # Rimuove i caratteri non validi per i nomi dei fogli Excel (/, \, ?, *, [, ], :)
                invalid_chars = '/\\?*[]:'
                for char in invalid_chars:
                    name = name.replace(char, '_')
                return name[:31] # Trunca a 31 caratteri, limite di Excel
                
            with pd.ExcelWriter(OUTPUT_FILE_INDIVIDUALE, engine="openpyxl") as writer:
                for cat, df in output_sheets.items():
                    if not df.empty:
                        sheet_name = sanitize_sheet_name(cat)
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
            risultato_msg += f"📄 Classifica Individuale generata: {OUTPUT_FILE_INDIVIDUALE}\n"

            if relay_sheets:
                with pd.ExcelWriter(OUTPUT_FILE_STAFFETTE, engine="openpyxl") as writer:
                    for cat, df in relay_sheets.items():
                        sheet_name = sanitize_sheet_name(cat)
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                risultato_msg += f"🏆 Classifica Staffette generata: {OUTPUT_FILE_STAFFETTE}\n"
            else:
                risultato_msg += "⚠️ Nessun dato sufficiente per generare la classifica delle staffette Excel.\n"

        except Exception as e:
            messagebox.showerror("Errore di Scrittura Excel", f"Errore durante la creazione dei file Excel. Errore: {e}")
            return
            
        # 6. GENERAZIONE PDF
        pdf_ind_ok = genera_pdf_individuale(output_sheets, print_time)
        if pdf_ind_ok:
            risultato_msg += f"📄 Classifica Individuale PDF generata: {OUTPUT_FILE_PDF_INDIVIDUALE}\n"
            
        # NON VIENE MODIFICATA genera_pdf_staffette come richiesto
        pdf_staff_ok = genera_pdf_staffette(relay_sheets, print_time)
        if pdf_staff_ok:
            risultato_msg += f"🏆 Classifica Staffette PDF generata: {OUTPUT_FILE_PDF_STAFFETTE}\n"


        if avvisi_staffette_generali:
             risultato_msg += "\nℹ️ Avvisi Elaborazione Staffette:\n" + "\n".join(avvisi_staffette_generali)


        messagebox.showinfo("Generazione Classifica Completata", risultato_msg)


if __name__ == "__main__":
    # Importo qui 'copy' per completezza, anche se è usato solo nella classe
    import copy
    app = App()
    app.mainloop()