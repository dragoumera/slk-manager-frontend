/* SLK Manager - Menu à GAUCHE */

* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

body {
  font-family: Arial, sans-serif;
  background: #f5f5f5;
}

.slk-manager {
  display: flex;
  flex-direction: row;
  min-height: 100vh;
}

/* SIDEBAR - GAUCHE */
.slk-sidebar,
nav,
[class*="sidebar"],
[class*="menu"] {
  width: 200px;
  background: #2c3e50;
  color: white;
  padding: 20px;
  border-right: 1px solid #1a252f;
  overflow-y: auto;
  flex-shrink: 0;
}

.slk-sidebar h3,
nav h3,
[class*="sidebar"] h3,
[class*="menu"] h3 {
  font-size: 12px;
  color: #bdc3c7;
  text-transform: uppercase;
  margin: 20px 0 10px 0;
  font-weight: 600;
}

.slk-sidebar button,
nav button,
[class*="sidebar"] button,
[class*="menu"] button {
  width: 100%;
  padding: 10px;
  background: #34495e;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 12px;
  text-align: left;
  margin-bottom: 5px;
  transition: all 0.2s ease;
}

.slk-sidebar button:hover,
nav button:hover,
[class*="sidebar"] button:hover,
[class*="menu"] button:hover {
  background: #455a64;
}

.slk-sidebar button.active,
nav button.active,
[class*="sidebar"] button.active,
[class*="menu"] button.active {
  background: #0066cc;
}

/* USER SECTION */
.user-section {
  margin-top: 30px;
  padding-top: 20px;
  border-top: 1px solid #34495e;
  font-size: 12px;
}

.user-section p {
  margin: 5px 0;
  color: #bdc3c7;
}

/* MAIN CONTENT - DROITE */
main,
.slk-content,
section {
  flex: 1;
  padding: 30px;
  overflow-y: auto;
  background: white;
}

header {
  background: linear-gradient(to right, #2c3e50, #34495e);
  color: white;
  padding: 20px 30px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.1);
}

header img {
  max-width: 50px;
  margin-bottom: 10px;
}

header h1 {
  margin: 5px 0;
  font-size: 20px;
}

header p {
  margin: 0;
  font-size: 12px;
  opacity: 0.9;
}

/* TABS */
.slk-tabs {
  display: flex;
  gap: 0;
  border-bottom: 2px solid #ddd;
  margin-bottom: 20px;
}

.tab-btn {
  padding: 12px 20px;
  background: transparent;
  border: none;
  border-bottom: 3px solid transparent;
  cursor: pointer;
  font-size: 13px;
  font-weight: 600;
  color: #666;
  transition: all 0.2s ease;
}

.tab-btn:hover {
  color: #0066cc;
  border-bottom-color: #0066cc;
}

.tab-btn.active {
  color: #0066cc;
  border-bottom-color: #0066cc;
}

/* SECTIONS */
section {
  display: none;
}

section.active {
  display: block;
}

section h2 {
  font-size: 24px;
  margin-bottom: 20px;
  border-bottom: 2px solid #0066cc;
  padding-bottom: 10px;
}

/* FORMS */
.slk-form-container {
  display: grid;
  grid-template-columns: 2fr 1fr;
  gap: 20px;
}

.slk-form,
.slk-list {
  background: #f9f9f9;
  padding: 20px;
  border-radius: 4px;
}

.slk-form h3,
.slk-list h3 {
  font-size: 14px;
  margin-bottom: 15px;
  font-weight: 600;
}

.form-row {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 15px;
  margin-bottom: 15px;
}

input,
textarea,
select {
  padding: 10px;
  border: 1px solid #ddd;
  border-radius: 4px;
  font-family: inherit;
  font-size: 13px;
  width: 100%;
}

input:focus,
textarea:focus,
select:focus {
  outline: none;
  border-color: #0066cc;
  box-shadow: 0 0 0 2px rgba(0, 102, 204, 0.1);
}

textarea {
  resize: vertical;
  min-height: 80px;
}

/* BUTTONS */
button[class*="btn"] {
  padding: 10px 20px;
  background: #0066cc;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 13px;
  font-weight: 600;
  transition: all 0.2s ease;
  margin-top: 10px;
}

button[class*="btn"]:hover {
  background: #0052a3;
}

button[class*="primary"] {
  width: 100%;
}

button[class*="print"],
button[class*="small"] {
  background: #27ae60;
  padding: 8px 12px;
  font-size: 12px;
  margin: 0;
  width: auto;
}

button[class*="print"]:hover,
button[class*="small"]:hover {
  background: #229954;
}

/* LIST */
.list-item {
  background: white;
  padding: 12px;
  margin-bottom: 8px;
  border-left: 3px solid #0066cc;
  border-radius: 4px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 15px;
}

.list-item p {
  margin: 3px 0;
  font-size: 12px;
}

.actions {
  display: flex;
  gap: 5px;
  margin-left: auto;
}

/* FOOTER */
footer {
  background: #2c3e50;
  color: white;
  text-align: center;
  padding: 20px;
  margin-top: 40px;
  font-size: 12px;
}

/* RESPONSIVE */
@media (max-width: 1024px) {
  .slk-manager {
    flex-direction: column;
  }

  .slk-sidebar,
  nav,
  [class*="sidebar"],
  [class*="menu"] {
    width: 100%;
    border-right: none;
    border-bottom: 1px solid #ddd;
  }

  .slk-form-container {
    grid-template-columns: 1fr;
  }
}

@media (max-width: 768px) {
  .form-row {
    grid-template-columns: 1fr;
  }
}

.hidden {
  display: none !important;
}
