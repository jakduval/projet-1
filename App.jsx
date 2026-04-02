import React from 'react';
import './App.css';

function App() {
    const handleDownload = () => {
        window.open('/api/download', '_blank');
    };

    return (
        <div className="container">
            <div className="icon">💰</div>
            <h1>Gestion de Trésorerie</h1>
            <p className="description">
                Fichier Excel complet avec formules automatiques pour suivre 
                et analyser votre trésorerie en toute simplicité.
            </p>
            
            <div className="features">
                <h3>📋 Contenu du fichier :</h3>
                <ul>
                    <li>Tableau de bord avec solde en temps réel</li>
                    <li>Suivi des mouvements (entrées/sorties)</li>
                    <li>Analyse par catégorie avec SUMIF</li>
                    <li>Prévisions sur 12 mois</li>
                    <li>Formules de calcul automatiques</li>
                    <li>Mise en forme professionnelle</li>
                </ul>
            </div>
            
            <button className="download-btn" onClick={handleDownload}>
                <span className="excel-icon">📥</span>
                Télécharger le fichier Excel
            </button>
            
            <p className="footer">
                Format .xlsx compatible Excel, Google Sheets, LibreOffice
            </p>
        </div>
    );
}

export default App;