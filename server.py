from fastapi import FastAPI
from fastapi.responses import Response
from fastapi.middleware.cors import CORSMiddleware
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from io import BytesIO
from datetime import datetime, timedelta

app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

@app.get("/")
def root():
    return {"message": "Fichier Excel de gestion de trésorerie prêt"}

@app.get("/api/download")
def download():
    wb = openpyxl.Workbook()
    
    # === STYLES ===
    header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    title_font = Font(bold=True, size=16, color="1E40AF")
    money_font_green = Font(bold=True, color="059669")
    money_font_red = Font(bold=True, color="DC2626")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center = Alignment(horizontal='center', vertical='center')
    
    # === FEUILLE 1: TABLEAU DE BORD ===
    ws1 = wb.active
    ws1.title = "Tableau de Bord"
    
    # Titre
    ws1.merge_cells('B2:E2')
    ws1['B2'] = "📊 TABLEAU DE BORD TRÉSORERIE"
    ws1['B2'].font = title_font
    ws1['B2'].alignment = center
    
    # Solde initial
    ws1['B4'] = "Solde Initial"
    ws1['C4'] = 10000
    ws1['C4'].number_format = '#,##0.00 €'
    ws1['C4'].font = Font(bold=True, size=14)
    
    # Résumé
    ws1['B6'] = "Total Entrées"
    ws1['C6'] = "=SUM(Mouvements!D:D)"
    ws1['C6'].number_format = '#,##0.00 €'
    ws1['C6'].font = money_font_green
    
    ws1['B7'] = "Total Sorties"
    ws1['C7'] = "=SUM(Mouvements!E:E)"
    ws1['C7'].number_format = '#,##0.00 €'
    ws1['C7'].font = money_font_red
    
    ws1['B9'] = "SOLDE ACTUEL"
    ws1['B9'].font = Font(bold=True, size=14)
    ws1['C9'] = "=C4+C6-C7"
    ws1['C9'].number_format = '#,##0.00 €'
    ws1['C9'].font = Font(bold=True, size=16, color="1E40AF")
    
    # Bordures et mise en forme
    for row in range(4, 10):
        for col in ['B', 'C']:
            ws1[f'{col}{row}'].border = border
    
    # Largeur colonnes
    ws1.column_dimensions['A'].width = 5
    ws1.column_dimensions['B'].width = 20
    ws1.column_dimensions['C'].width = 18
    
    # === FEUILLE 2: MOUVEMENTS ===
    ws2 = wb.create_sheet("Mouvements")
    
    # En-têtes
    headers = ["Date", "Description", "Catégorie", "Entrées (+)", "Sorties (-)", "Solde"]
    for col, header in enumerate(headers, 1):
        cell = ws2.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center
    
    # Données exemple
    today = datetime.now()
    mouvements = [
        (today - timedelta(days=30), "Solde reporté", "Report", 10000, 0),
        (today - timedelta(days=28), "Vente produit A", "Ventes", 2500, 0),
        (today - timedelta(days=25), "Loyer bureau", "Loyer", 0, 1200),
        (today - timedelta(days=22), "Facture client B", "Ventes", 3800, 0),
        (today - timedelta(days=20), "Salaires", "Salaires", 0, 4500),
        (today - timedelta(days=18), "Achat fournitures", "Fournitures", 0, 350),
        (today - timedelta(days=15), "Vente produit C", "Ventes", 1800, 0),
        (today - timedelta(days=12), "Electricité", "Charges", 0, 280),
        (today - timedelta(days=10), "Prestation service", "Services", 4200, 0),
        (today - timedelta(days=8), "Assurance", "Assurance", 0, 450),
        (today - timedelta(days=5), "Remboursement TVA", "TVA", 1200, 0),
        (today - timedelta(days=3), "Internet/Téléphone", "Charges", 0, 120),
        (today - timedelta(days=1), "Vente en ligne", "Ventes", 950, 0),
        (today, "Frais bancaires", "Banque", 0, 45),
    ]
    
    for row_idx, (date, desc, cat, entree, sortie) in enumerate(mouvements, 2):
        ws2.cell(row=row_idx, column=1, value=date).number_format = 'DD/MM/YYYY'
        ws2.cell(row=row_idx, column=2, value=desc)
        ws2.cell(row=row_idx, column=3, value=cat)
        
        cell_entree = ws2.cell(row=row_idx, column=4, value=entree if entree > 0 else None)
        cell_entree.number_format = '#,##0.00 €'
        if entree > 0:
            cell_entree.font = money_font_green
        
        cell_sortie = ws2.cell(row=row_idx, column=5, value=sortie if sortie > 0 else None)
        cell_sortie.number_format = '#,##0.00 €'
        if sortie > 0:
            cell_sortie.font = money_font_red
        
        # Formule solde cumulé
        if row_idx == 2:
            ws2.cell(row=row_idx, column=6, value=f"=D{row_idx}-E{row_idx}")
        else:
            ws2.cell(row=row_idx, column=6, value=f"=F{row_idx-1}+D{row_idx}-E{row_idx}")
        ws2.cell(row=row_idx, column=6).number_format = '#,##0.00 €'
        ws2.cell(row=row_idx, column=6).font = Font(bold=True)
        
        # Bordures
        for col in range(1, 7):
            ws2.cell(row=row_idx, column=col).border = border
    
    # Largeur colonnes
    ws2.column_dimensions['A'].width = 14
    ws2.column_dimensions['B'].width = 25
    ws2.column_dimensions['C'].width = 15
    ws2.column_dimensions['D'].width = 15
    ws2.column_dimensions['E'].width = 15
    ws2.column_dimensions['F'].width = 15
    
    # === FEUILLE 3: CATÉGORIES ===
    ws3 = wb.create_sheet("Analyse par Catégorie")
    
    ws3.merge_cells('B2:D2')
    ws3['B2'] = "📈 ANALYSE PAR CATÉGORIE"
    ws3['B2'].font = title_font
    ws3['B2'].alignment = center
    
    # En-têtes
    headers_cat = ["Catégorie", "Total Entrées", "Total Sorties"]
    for col, header in enumerate(headers_cat, 2):
        cell = ws3.cell(row=4, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center
    
    categories = ["Ventes", "Services", "Loyer", "Salaires", "Fournitures", "Charges", "Assurance", "TVA", "Banque", "Report"]
    for row_idx, cat in enumerate(categories, 5):
        ws3.cell(row=row_idx, column=2, value=cat).border = border
        
        # SUMIF pour entrées
        cell_e = ws3.cell(row=row_idx, column=3, 
                         value=f'=SUMIF(Mouvements!C:C,B{row_idx},Mouvements!D:D)')
        cell_e.number_format = '#,##0.00 €'
        cell_e.border = border
        cell_e.font = money_font_green
        
        # SUMIF pour sorties
        cell_s = ws3.cell(row=row_idx, column=4,
                         value=f'=SUMIF(Mouvements!C:C,B{row_idx},Mouvements!E:E)')
        cell_s.number_format = '#,##0.00 €'
        cell_s.border = border
        cell_s.font = money_font_red
    
    # Largeur colonnes
    ws3.column_dimensions['A'].width = 5
    ws3.column_dimensions['B'].width = 18
    ws3.column_dimensions['C'].width = 18
    ws3.column_dimensions['D'].width = 18
    
    # === FEUILLE 4: PRÉVISIONS ===
    ws4 = wb.create_sheet("Prévisions")
    
    ws4.merge_cells('B2:E2')
    ws4['B2'] = "🔮 PRÉVISIONS TRÉSORERIE"
    ws4['B2'].font = title_font
    ws4['B2'].alignment = center
    
    # En-têtes
    headers_prev = ["Mois", "Entrées Prévues", "Sorties Prévues", "Solde Prévu"]
    for col, header in enumerate(headers_prev, 2):
        cell = ws4.cell(row=4, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center
    
    mois = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
            "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
    
    for row_idx, m in enumerate(mois, 5):
        ws4.cell(row=row_idx, column=2, value=m).border = border
        
        cell_e = ws4.cell(row=row_idx, column=3, value=0)
        cell_e.number_format = '#,##0.00 €'
        cell_e.border = border
        
        cell_s = ws4.cell(row=row_idx, column=4, value=0)
        cell_s.number_format = '#,##0.00 €'
        cell_s.border = border
        
        # Formule solde
        if row_idx == 5:
            ws4.cell(row=row_idx, column=5, value=f"='Tableau de Bord'!C9+C{row_idx}-D{row_idx}")
        else:
            ws4.cell(row=row_idx, column=5, value=f"=E{row_idx-1}+C{row_idx}-D{row_idx}")
        ws4.cell(row=row_idx, column=5).number_format = '#,##0.00 €'
        ws4.cell(row=row_idx, column=5).font = Font(bold=True)
        ws4.cell(row=row_idx, column=5).border = border
    
    # Largeur colonnes
    ws4.column_dimensions['A'].width = 5
    ws4.column_dimensions['B'].width = 15
    ws4.column_dimensions['C'].width = 18
    ws4.column_dimensions['D'].width = 18
    ws4.column_dimensions['E'].width = 18
    
    # Sauvegarder
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    return Response(
        content=buffer.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=gestion_tresorerie.xlsx"}
    )