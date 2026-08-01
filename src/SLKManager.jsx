import React, { useState } from 'react';
import './SLKManager.css';

/**
 * SLK CLIM Manager
 * Gestionnaire de devis, factures et PV
 * Format d'impression: A4 sans marges
 */

const SLKManager = () => {
  const [activeTab, setActiveTab] = useState('devis');
  const [devis, setDevis] = useState([]);
  const [factures, setFactures] = useState([]);
  const [pvs, setPvs] = useState([]);
  const [formData, setFormData] = useState({
    client: '',
    email: '',
    telephone: '',
    adresse: '',
    entreprise: '',
    dateDebut: new Date().toISOString().split('T')[0],
    dateFin: new Date().toISOString().split('T')[0],
    description: '',
    montant: '',
    articles: [],
    notes: ''
  });

  // Logo path - À PLACER DANS LE DOSSIER public/images/
  // Utilisation: /images/logo-srv-gaine.png
  const LOGO_PATH = '/images/logo-srv-gaine.png';

  const handleFormChange = (e) => {
    const { name, value } = e.target;
    setFormData(prev => ({
      ...prev,
      [name]: value
    }));
  };

  const generateDevis = () => {
    if (!formData.client || !formData.montant) {
      alert('Veuillez remplir tous les champs obligatoires');
      return;
    }

    const newDevis = {
      id: `DEV-${Date.now()}`,
      ...formData,
      date: new Date().toLocaleDateString('fr-FR'),
      statut: 'En attente'
    };

    setDevis([...devis, newDevis]);
    resetForm();
    alert('Devis créé avec succès!');
  };

  const generateFacture = () => {
    if (!formData.client || !formData.montant) {
      alert('Veuillez remplir tous les champs obligatoires');
      return;
    }

    const newFacture = {
      id: `FAC-${Date.now()}`,
      ...formData,
      date: new Date().toLocaleDateString('fr-FR'),
      statut: 'Payée'
    };

    setFactures([...factures, newFacture]);
    resetForm();
    alert('Facture créée avec succès!');
  };

  const generatePV = () => {
    if (!formData.client || !formData.description) {
      alert('Veuillez remplir tous les champs obligatoires');
      return;
    }

    const newPV = {
      id: `PV-${Date.now()}`,
      ...formData,
      date: new Date().toLocaleDateString('fr-FR'),
      statut: 'Validé'
    };

    setPvs([...pvs, newPV]);
    resetForm();
    alert('PV créé avec succès!');
  };

  const resetForm = () => {
    setFormData({
      client: '',
      email: '',
      telephone: '',
      adresse: '',
      entreprise: '',
      dateDebut: new Date().toISOString().split('T')[0],
      dateFin: new Date().toISOString().split('T')[0],
      description: '',
      montant: '',
      articles: [],
      notes: ''
    });
  };

  const imprimerDevis = (devisItem) => {
    window.open(`/print/devis/${devisItem.id}`, '_blank');
  };

  const imprimerFacture = (factureItem) => {
    window.open(`/print/facture/${factureItem.id}`, '_blank');
  };

  const imprimerPV = (pvItem) => {
    window.open(`/print/pv/${pvItem.id}`, '_blank');
  };

  return (
    <div className="slk-manager">
      {/* HEADER */}
      <header className="slk-header">
        <div className="slk-header-content">
          <div className="slk-logo-section">
            {/* LOGO - À METTRE DANS: public/images/logo-srv-gaine.png */}
            <img src={LOGO_PATH} alt="SRV GAINE Logo" className="slk-logo" />
            <div>
              <h1>SLK CLIM</h1>
              <p>Climatisation & Ventilation</p>
            </div>
          </div>
          <div className="slk-header-info">
            <p><strong>📍</strong> 123 Rue de la Ventilation, 75000 Paris</p>
            <p><strong>📞</strong> +33 1 23 45 67 89 | <strong>📧</strong> contact@slkclim.com</p>
          </div>
        </div>
      </header>

      {/* NAVIGATION TABS */}
      <nav className="slk-tabs">
        <button 
          className={`tab-btn ${activeTab === 'devis' ? 'active' : ''}`}
          onClick={() => setActiveTab('devis')}
        >
          📄 DEVIS
        </button>
        <button 
          className={`tab-btn ${activeTab === 'factures' ? 'active' : ''}`}
          onClick={() => setActiveTab('factures')}
        >
          💰 FACTURES
        </button>
        <button 
          className={`tab-btn ${activeTab === 'pvs' ? 'active' : ''}`}
          onClick={() => setActiveTab('pvs')}
        >
          ✅ PV
        </button>
      </nav>

      {/* MAIN CONTENT */}
      <main className="slk-content">
        {/* SECTION DEVIS */}
        {activeTab === 'devis' && (
          <section className="slk-section">
            <h2>Créer un Devis</h2>
            <div className="slk-form-container">
              <div className="slk-form">
                <h3>Informations Client</h3>
                <div className="form-row">
                  <input
                    type="text"
                    name="client"
                    placeholder="Nom du client *"
                    value={formData.client}
                    onChange={handleFormChange}
                    required
                  />
                  <input
                    type="text"
                    name="entreprise"
                    placeholder="Entreprise"
                    value={formData.entreprise}
                    onChange={handleFormChange}
                  />
                </div>

                <div className="form-row">
                  <input
                    type="email"
                    name="email"
                    placeholder="Email"
                    value={formData.email}
                    onChange={handleFormChange}
                  />
                  <input
                    type="tel"
                    name="telephone"
                    placeholder="Téléphone"
                    value={formData.telephone}
                    onChange={handleFormChange}
                  />
                </div>

                <input
                  type="text"
                  name="adresse"
                  placeholder="Adresse"
                  value={formData.adresse}
                  onChange={handleFormChange}
                  style={{ width: '100%' }}
                />

                <h3 style={{ marginTop: '20px' }}>Détails du Devis</h3>
                <textarea
                  name="description"
                  placeholder="Description du travail *"
                  value={formData.description}
                  onChange={handleFormChange}
                  rows="4"
                ></textarea>

                <div className="form-row">
                  <input
                    type="number"
                    name="montant"
                    placeholder="Montant HT (€) *"
                    value={formData.montant}
                    onChange={handleFormChange}
                    step="0.01"
                  />
                  <input
                    type="date"
                    name="dateDebut"
                    value={formData.dateDebut}
                    onChange={handleFormChange}
                  />
                </div>

                <textarea
                  name="notes"
                  placeholder="Notes supplémentaires"
                  value={formData.notes}
                  onChange={handleFormChange}
                  rows="2"
                ></textarea>

                <button className="btn-primary" onClick={generateDevis}>
                  💾 Créer le Devis
                </button>
              </div>

              <div className="slk-list">
                <h3>Devis Créés</h3>
                {devis.length === 0 ? (
                  <p style={{ color: '#999' }}>Aucun devis pour le moment</p>
                ) : (
                  devis.map(d => (
                    <div key={d.id} className="list-item">
                      <div>
                        <p><strong>{d.id}</strong></p>
                        <p>{d.client} - {d.montant}€</p>
                        <p style={{ fontSize: '12px', color: '#666' }}>{d.date}</p>
                      </div>
                      <div className="actions">
                        <button className="btn-small btn-print" onClick={() => imprimerDevis(d)}>
                          🖨️ Imprimer
                        </button>
                      </div>
                    </div>
                  ))
                )}
              </div>
            </div>
          </section>
        )}

        {/* SECTION FACTURES */}
        {activeTab === 'factures' && (
          <section className="slk-section">
            <h2>Créer une Facture</h2>
            <div className="slk-form-container">
              <div className="slk-form">
                <h3>Informations Client</h3>
                <div className="form-row">
                  <input
                    type="text"
                    name="client"
                    placeholder="Nom du client *"
                    value={formData.client}
                    onChange={handleFormChange}
                    required
                  />
                  <input
                    type="text"
                    name="entreprise"
                    placeholder="Entreprise"
                    value={formData.entreprise}
                    onChange={handleFormChange}
                  />
                </div>

                <div className="form-row">
                  <input
                    type="email"
                    name="email"
                    placeholder="Email"
                    value={formData.email}
                    onChange={handleFormChange}
                  />
                  <input
                    type="tel"
                    name="telephone"
                    placeholder="Téléphone"
                    value={formData.telephone}
                    onChange={handleFormChange}
                  />
                </div>

                <input
                  type="text"
                  name="adresse"
                  placeholder="Adresse"
                  value={formData.adresse}
                  onChange={handleFormChange}
                  style={{ width: '100%' }}
                />

                <h3 style={{ marginTop: '20px' }}>Détails de la Facture</h3>
                <textarea
                  name="description"
                  placeholder="Description du travail facturé *"
                  value={formData.description}
                  onChange={handleFormChange}
                  rows="4"
                ></textarea>

                <div className="form-row">
                  <input
                    type="number"
                    name="montant"
                    placeholder="Montant TTC (€) *"
                    value={formData.montant}
                    onChange={handleFormChange}
                    step="0.01"
                  />
                  <input
                    type="date"
                    name="dateDebut"
                    value={formData.dateDebut}
                    onChange={handleFormChange}
                  />
                </div>

                <textarea
                  name="notes"
                  placeholder="Conditions de paiement"
                  value={formData.notes}
                  onChange={handleFormChange}
                  rows="2"
                ></textarea>

                <button className="btn-primary" onClick={generateFacture}>
                  💾 Créer la Facture
                </button>
              </div>

              <div className="slk-list">
                <h3>Factures Créées</h3>
                {factures.length === 0 ? (
                  <p style={{ color: '#999' }}>Aucune facture pour le moment</p>
                ) : (
                  factures.map(f => (
                    <div key={f.id} className="list-item">
                      <div>
                        <p><strong>{f.id}</strong></p>
                        <p>{f.client} - {f.montant}€</p>
                        <p style={{ fontSize: '12px', color: '#666' }}>{f.date}</p>
                      </div>
                      <div className="actions">
                        <button className="btn-small btn-print" onClick={() => imprimerFacture(f)}>
                          🖨️ Imprimer
                        </button>
                      </div>
                    </div>
                  ))
                )}
              </div>
            </div>
          </section>
        )}

        {/* SECTION PV */}
        {activeTab === 'pvs' && (
          <section className="slk-section">
            <h2>Créer un Procès-Verbal</h2>
            <div className="slk-form-container">
              <div className="slk-form">
                <h3>Informations Client</h3>
                <div className="form-row">
                  <input
                    type="text"
                    name="client"
                    placeholder="Nom du client *"
                    value={formData.client}
                    onChange={handleFormChange}
                    required
                  />
                  <input
                    type="text"
                    name="entreprise"
                    placeholder="Entreprise"
                    value={formData.entreprise}
                    onChange={handleFormChange}
                  />
                </div>

                <div className="form-row">
                  <input
                    type="email"
                    name="email"
                    placeholder="Email"
                    value={formData.email}
                    onChange={handleFormChange}
                  />
                  <input
                    type="tel"
                    name="telephone"
                    placeholder="Téléphone"
                    value={formData.telephone}
                    onChange={handleFormChange}
                  />
                </div>

                <input
                  type="text"
                  name="adresse"
                  placeholder="Adresse"
                  value={formData.adresse}
                  onChange={handleFormChange}
                  style={{ width: '100%' }}
                />

                <h3 style={{ marginTop: '20px' }}>Détails du PV</h3>
                <div className="form-row">
                  <input
                    type="date"
                    name="dateDebut"
                    label="Date de début"
                    value={formData.dateDebut}
                    onChange={handleFormChange}
                  />
                  <input
                    type="date"
                    name="dateFin"
                    label="Date de fin"
                    value={formData.dateFin}
                    onChange={handleFormChange}
                  />
                </div>

                <textarea
                  name="description"
                  placeholder="Description des interventions / Observations *"
                  value={formData.description}
                  onChange={handleFormChange}
                  rows="5"
                ></textarea>

                <textarea
                  name="notes"
                  placeholder="Remarques et signatures"
                  value={formData.notes}
                  onChange={handleFormChange}
                  rows="3"
                ></textarea>

                <button className="btn-primary" onClick={generatePV}>
                  💾 Créer le PV
                </button>
              </div>

              <div className="slk-list">
                <h3>PV Créés</h3>
                {pvs.length === 0 ? (
                  <p style={{ color: '#999' }}>Aucun PV pour le moment</p>
                ) : (
                  pvs.map(pv => (
                    <div key={pv.id} className="list-item">
                      <div>
                        <p><strong>{pv.id}</strong></p>
                        <p>{pv.client}</p>
                        <p style={{ fontSize: '12px', color: '#666' }}>
                          {pv.dateDebut} → {pv.dateFin}
                        </p>
                      </div>
                      <div className="actions">
                        <button className="btn-small btn-print" onClick={() => imprimerPV(pv)}>
                          🖨️ Imprimer
                        </button>
                      </div>
                    </div>
                  ))
                )}
              </div>
            </div>
          </section>
        )}
      </main>

      {/* FOOTER */}
      <footer className="slk-footer">
        <p>&copy; 2026 SLK CLIM. Tous droits réservés.</p>
        <p>Conception et réalisation par <strong>INFORMAINT SARL</strong></p>
      </footer>
    </div>
  );
};

export default SLKManager;
