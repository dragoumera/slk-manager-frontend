import React, { useState } from 'react';
import SLKManager from './SLKManager.jsx';
import './App.css';

export default function App() {
  const [isLoggedIn, setIsLoggedIn] = useState(false);

  const handleLogin = async (e) => {
    e.preventDefault();
    const email = e.target.email.value;
    const password = e.target.password.value;

    try {
      const response = await fetch('https://slk-manager-api.onrender.com/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, password })
      });

      const data = await response.json();

      if (response.ok && data.token) {
        localStorage.setItem('token', data.token);
        setIsLoggedIn(true);
      } else {
        if (email === 'dragoumera@gmail.com' && password === 'S@voir12345') {
          setIsLoggedIn(true);
        } else {
          alert('❌ Email ou mot de passe incorrect');
        }
      }
    } catch (error) {
      if (email === 'dragoumera@gmail.com' && password === 'S@voir12345') {
        setIsLoggedIn(true);
      } else {
        alert('❌ Erreur: ' + error.message);
      }
    }
  };

  if (isLoggedIn) {
    return (
      <div>
        <button
          onClick={() => setIsLoggedIn(false)}
          style={{
            position: 'fixed',
            top: '20px',
            right: '20px',
            padding: '10px 20px',
            background: '#e74c3c',
            color: 'white',
            border: 'none',
            borderRadius: '6px',
            cursor: 'pointer',
            fontWeight: 'bold',
            zIndex: 200
          }}
        >
          ← Retour à l'accueil
        </button>
        <SLKManager />
      </div>
    );
  }

  return (
    <div className="login-page">
      <div className="login-container">
        <div className="login-box">
          <img src="/images/logo-srv-gaine.png" alt="Logo" className="login-logo" />
          <h1>SLK CLIM</h1>
          <p>Espace Administrateur</p>
          
          <form onSubmit={handleLogin}>
            <div className="form-group">
              <label>Email</label>
              <input type="email" name="email" required />
            </div>
            <div className="form-group">
              <label>Mot de Passe</label>
              <input type="password" name="password" required />
            </div>
            <button type="submit" className="btn-login">Se Connecter</button>
          </form>

          <p className="login-info">
            Compte test: dragoumera@gmail.com / S@voir12345
          </p>
        </div>
      </div>
    </div>
  );
}
