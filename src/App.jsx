import React, { useState } from 'react';
import './App.css';

export default function App() {
  const [currentSection, setCurrentSection] = useState('accueil');

  const showSection = (sectionId) => {
    setCurrentSection(sectionId);
    window.scrollTo(0, 0);
  };

  const handleLogin = (e) => {
    e.preventDefault();
    const email = e.target.email.value;
    const password = e.target.password.value;

    if (email === 'dragoumera@gmail.com' && password === 'S@voir12345') {
      alert('✅ Connecté en tant qu\'Administrateur!');
      window.location.href = '/manager';
    } else {
      alert('❌ Email ou mot de passe incorrect!');
    }
  };

  return (
    <div className="app">
      {/* HEADER */}
      <header>
        <div className="header-container">
          <div className="logo-section">
            <div className="logo">
              <img src="/images/logo-srv-gaine.png" alt="SRV GAINE Logo" />
            </div>
            <div className="company-info">
              <h1>SLK CLIM</h1>
              <p>Climatisation & Ventilation</p>
            </div>
          </div>
          <div style={{ textAlign: 'right', fontSize: '13px' }}>
            <p>📍 8 Av. Roland Moreno, 95740 Frépillon</p>
            <p>📞 01 87 63 23 76 | 📱 06 11 12 00 61 | 📧 slk.clim@yahoo.fr</p>
          </div>
          <nav>
            <ul>
              <li><a onClick={() => showSection('accueil')}>Accueil</a></li>
              <li><a onClick={() => showSection('apropos')}>À Propos</a></li>
              <li><a onClick={() => showSection('services')}>Services</a></li>
              <li><a onClick={() => showSection('produits')}>Produits</a></li>
              <li><a onClick={() => showSection('projets')}>Projets</a></li>
              <li><a onClick={() => showSection('contact')}>Contact</a></li>
              <li><a onClick={() => showSection('connexion')} className="connexion">Connexion</a></li>
            </ul>
          </nav>
        </div>
      </header>

      {/* ACCUEIL */}
      {currentSection === 'accueil' && (
        <section id="accueil" className="hero">
          <div className="hero-content">
            <h2>Climatisation & Ventilation Professionnelle</h2>
            <p>Solutions innovantes pour votre confort thermique</p>
            <div className="cta-buttons">
              <button className="btn btn-primary" onClick={() => showSection('contact')}>Nous Contacter</button>
              <button className="btn" onClick={() => showSection('services')}>En savoir plus</button>
            </div>
          </div>
        </section>
      )}

      {/* À PROPOS */}
      {currentSection === 'apropos' && (
        <section id="apropos">
          <h2>À Propos</h2>
          <div className="cards-grid">
            <div className="card">
              <h3>📍 20+ Ans d'Expérience</h3>
              <p>SLK CLIM intervient depuis plus de 20 ans dans le domaine de la climatisation et ventilation.</p>
            </div>
            <div className="card">
              <h3>👨‍🔧 50+ Techniciens</h3>
              <p>Une équipe expérimentée et certifiée pour tous vos besoins de climatisation.</p>
            </div>
            <div className="card">
              <h3>✅ 500+ Projets</h3>
              <p>Plus de 500 projets réussis auprès de clients résidentiels et commerciaux.</p>
            </div>
            <div className="card">
              <h3>🌍 1000+ Clients</h3>
              <p>La confiance de plus de 1000 clients satisfaits en Île-de-France.</p>
            </div>
          </div>
        </section>
      )}

      {/* SERVICES */}
      {currentSection === 'services' && (
        <section id="services">
          <h2>Nos Services</h2>
          <div className="cards-grid">
            <div className="card">
              <h3>🔧 Installation</h3>
              <p>Installation complète de systèmes de climatisation et ventilation sur mesure.</p>
            </div>
            <div className="card">
              <h3>🛠️ Maintenance</h3>
              <p>Maintenance régulière pour garantir le bon fonctionnement de vos équipements.</p>
            </div>
            <div className="card">
              <h3>🔍 Diagnostic</h3>
              <p>Diagnostic complet de votre système pour optimiser votre confort.</p>
            </div>
            <div className="card">
              <h3>♻️ Éco-Responsable</h3>
              <p>Solutions écologiques et économes en énergie pour votre bien-être.</p>
            </div>
            <div className="card">
              <h3>📞 Support 24/7</h3>
              <p>Service client disponible pour toutes vos questions et urgences.</p>
            </div>
            <div className="card">
              <h3>📋 Garantie</h3>
              <p>Tous nos travaux sont garantis avec des devis transparents.</p>
            </div>
          </div>
        </section>
      )}

      {/* PRODUITS */}
      {currentSection === 'produits' && (
        <section id="produits">
          <h2>Nos Produits</h2>
          <div className="cards-grid">
            <div className="card">
              <h3>🔴 Gaines Rondes</h3>
              <p>Gaines de ventilation rondes haute performance pour installations efficaces.</p>
            </div>
            <div className="card">
              <h3>⬜ Gaines Carrées</h3>
              <p>Gaines carrées adaptées aux configurations d'espace restreint.</p>
            </div>
            <div className="card">
              <h3>⚙️ Accessoires</h3>
              <p>Accessoires complets pour climatisation et ventilation.</p>
            </div>
            <div className="card">
              <h3>🔩 Pièces de Rechange</h3>
              <p>Toutes les pièces de rechange pour maintenance et réparation.</p>
            </div>
            <div className="card">
              <h3>❄️ Climatiseurs</h3>
              <p>Climatiseurs dernière génération avec technologie silencieuse.</p>
            </div>
            <div className="card">
              <h3>🌬️ Filtres</h3>
              <p>Filtres haute efficacité pour une meilleure qualité d'air.</p>
            </div>
          </div>
        </section>
      )}

      {/* PROJETS */}
      {currentSection === 'projets' && (
        <section id="projets">
          <h2>Nos Projets</h2>
          <div className="cards-grid">
            <div className="card">
              <h3>🏢 Centre Commercial</h3>
              <p>Installation climatisation pour 5000m² - Janvier 2024</p>
            </div>
            <div className="card">
              <h3>🏥 Hôpital</h3>
              <p>Système de ventilation médical pour blocs opératoires - Mars 2024</p>
            </div>
            <div className="card">
              <h3>🏭 Usine de Production</h3>
              <p>Climatisation industrielle pour maintien température - Février 2024</p>
            </div>
            <div className="card">
              <h3>🏠 Résidence Luxe</h3>
              <p>Installation climatisation silencieuse 50 appartements - Avril 2024</p>
            </div>
            <div className="card">
              <h3>🛒 Supermarché</h3>
              <p>Système frigorifique et climatisation pour stockage - Mai 2024</p>
            </div>
            <div className="card">
              <h3>🏢 Bureau Moderne</h3>
              <p>Climatisation écologique pour immeuble 10 étages - Juin 2024</p>
            </div>
          </div>
        </section>
      )}

      {/* CONTACT */}
      {currentSection === 'contact' && (
        <section id="contact">
          <h2>Nous Contacter</h2>
          <div className="cards-grid">
            <div className="card">
              <h3>📍 Adresse</h3>
              <p>8 Avenue Roland Moreno<br />Bâtiment B 2<br />95740 Frépillon</p>
            </div>
            <div className="card">
              <h3>📞 Téléphone</h3>
              <p>+33 1 87 63 23 76<br />📱 +33 6 11 12 00 61<br />Lun-Ven: 8h-18h</p>
            </div>
            <div className="card">
              <h3>📧 Email</h3>
              <p>slk.clim@yahoo.fr<br />Réponse en 24h</p>
            </div>
          </div>
        </section>
      )}

      {/* CONNEXION */}
      {currentSection === 'connexion' && (
        <section id="connexion">
          <h2>Espace Admin - Connexion</h2>
          <form onSubmit={handleLogin} style={{ maxWidth: '400px', margin: '40px auto' }}>
            <div className="form-group">
              <label>Email *</label>
              <input type="email" name="email" required />
            </div>
            <div className="form-group">
              <label>Mot de Passe *</label>
              <input type="password" name="password" required />
            </div>
            <div className="form-group">
              <button type="submit">Se Connecter</button>
            </div>
            <p style={{ textAlign: 'center', fontSize: '12px', color: '#666', marginTop: '15px' }}>
              <strong>Admin par défaut:</strong><br />
              Email: dragoumera@gmail.com<br />
              Mot de passe: S@voir12345
            </p>
          </form>
        </section>
      )}

      {/* FOOTER */}
      <footer>
        <p>&copy; 2026 SLK CLIM. Tous droits réservés.</p>
        <p>Conception et réalisation par <strong>INFORMAINT SARL</strong></p>
      </footer>
    </div>
  );
}

