const GAS_URL = "VOTRE_URL_SCRIPT_ICI";

document.getElementById('labelForm').onsubmit = async (e) => {
    e.preventDefault();
    const btn = document.getElementById('btnText');
    const fileInput = document.getElementById('excelFile');
    
    if(!fileInput.files[0]) return alert("Veuillez charger un fichier Excel.");
    
    btn.innerText = "TRAITEMENT EN COURS...";
    
    const reader = new FileReader();
    reader.onload = async (event) => {
        try {
            const wb = XLSX.read(new Uint8Array(event.target.result), {type: 'array'});
            const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {raw:false});
            const codes = data.map(r => r.Code || r.code || r.CODE).filter(c => c);

            const payload = {
                rows: codes,
                options: {
                    type: document.getElementById('type').value,
                    preset: document.getElementById('preset').value,
                    showText: true,
                    textPos: document.getElementById('textPos').value,
                    email: document.getElementById('userEmail').value,
                    // Valeurs personnalisées
                    custW: document.getElementById('custW').value,
                    custH: document.getElementById('custH').value,
                    custCols: document.getElementById('custCols').value,
                    custRows: document.getElementById('custRows').value
                }
            };

            await fetch(GAS_URL, {
                method: 'POST',
                mode: 'no-cors',
                body: JSON.stringify(payload)
            });

            btn.innerText = "GÉNÉRER LE PDF";
            document.getElementById('result').classList.remove('hidden');
            alert("Terminé ! Le PDF sera disponible sur Drive et par email dans quelques instants.");
        } catch (err) {
            alert("Erreur technique : " + err.message);
            btn.innerText = "GÉNÉRER LE PDF";
        }
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
};
