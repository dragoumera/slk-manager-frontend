* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

body {
  font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
  background: #f5f5f5;
}

/* LOGIN PAGE */
.login-page {
  display: flex;
  justify-content: center;
  align-items: center;
  min-height: 100vh;
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
}

.login-container {
  width: 100%;
  max-width: 400px;
  padding: 20px;
}

.login-box {
  background: white;
  padding: 40px;
  border-radius: 12px;
  box-shadow: 0 10px 40px rgba(0,0,0,0.2);
  text-align: center;
}

.login-logo {
  width: 80px;
  height: 80px;
  margin-bottom: 20px;
  border-radius: 8px;
  background: #f0f0f0;
  padding: 10px;
}

.login-box h1 {
  font-size: 28px;
  color: #2c3e50;
  margin-bottom: 8px;
}

.login-box p {
  color: #7f8c8d;
  margin-bottom: 30px;
  font-size: 14px;
}

.form-group {
  margin-bottom: 20px;
  text-align: left;
}

.form-group label {
  display: block;
  margin-bottom: 8px;
  font-weight: 600;
  color: #2c3e50;
  font-size: 14px;
}

.form-group input {
  width: 100%;
  padding: 12px;
  border: 2px solid #ecf0f1;
  border-radius: 6px;
  font-size: 14px;
  transition: all 0.3s ease;
  font-family: inherit;
}

.form-group input:focus {
  outline: none;
  border-color: #667eea;
  box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
}

.btn-login {
  width: 100%;
  padding: 12px;
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
  border: none;
  border-radius: 6px;
  font-weight: 600;
  font-size: 15px;
  cursor: pointer;
  transition: all 0.3s ease;
  text-transform: uppercase;
  letter-spacing: 0.5px;
}

.btn-login:hover {
  transform: translateY(-2px);
  box-shadow: 0 8px 20px rgba(102, 126, 234, 0.3);
}

.login-info {
  margin-top: 20px;
  font-size: 12px;
  color: #7f8c8d;
  background: #f9f9f9;
  padding: 12px;
  border-radius: 6px;
}

/* RESPONSIVE */
@media (max-width: 480px) {
  .login-box {
    padding: 30px 20px;
  }

  .login-logo {
    width: 60px;
    height: 60px;
  }

  .login-box h1 {
    font-size: 24px;
  }
}
