import React from 'react';
import { useNavigate } from 'react-router-dom';
import SLKManager from './SLKManager.jsx';
import './SLKManager.css';

export default function SLKManagerPage() {
  const navigate = useNavigate();

  return (
    <div style={{ position: 'relative' }}>
      <button
        onClick={() => navigate('/')}
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
        ← Retour au Site
      </button>
      <SLKManager />
    </div>
  );
}
