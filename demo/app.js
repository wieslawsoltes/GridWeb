import './app-04.js';
import './pivot-ui.js';
import './editing-ui.js';
import './reference-ui.js';
// Commands that activate a newly created sheet must also update the visible tab state.
document.getElementById('grid').addEventListener('selection-change',()=>{
 const name=window.gridweb.grid.Sheet.Name;
 for(const tab of document.querySelectorAll('#sheet-tabs [role="tab"]'))tab.setAttribute('aria-selected',String(tab.textContent===name));
});
