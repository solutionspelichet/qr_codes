const GAS_URL = "https://script.google.com/macros/s/AKfycbxDyfl3NhpEbFEL7zCdQgRCQrIutFqRwSnd1gGUOosaWIiIx5PdNQJbZPuSu_7mneUq/exec";

document.getElementById('labelForm').onsubmit = async (e) => {
    e.preventDefault();
    const btn = document.getElementById('btnText');
    const fileInput = document.getElementById('excelFile');
    
    if(!fileInput.files[0]) return alert("Sélectionnez un fichier");
    
    btn.innerText = "TRAITEMENT EN COURS...";
    
    const reader = new FileReader();
    reader.onload = async (event) => {
        const wb = XLSX.read(new Uint8Array(event.target.result), {type: 'array'});
        const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {raw:false});
        const codes = data.map(r => r.Code || r.code || r.CODE).filter(c => c);

        const payload = {
            rows: codes,
            options: {
                type: document.getElementById('type').value,
                preset: document.getElementById('preset').value,
                showText: document.getElementById('showText').value === "true",
                textPos: document.getElementById('textPos').value,
                email: document.getElementById('userEmail').value
            }
        };

        try {
            // L'appel fetch avec mode no-cors ne permet pas de lire la réponse
            // Mais déclenche bien l'exécution côté Google.
            await fetch(GAS_URL, {
                method: 'POST',
                mode: 'no-cors',
                body: JSON.stringify(payload)
            });

            btn.innerText = "GÉNÉRER PDF";
            document.getElementById('result').classList.remove('hidden');
            document.getElementById('driveLink').href = "https://drive.google.com/";
            alert("Requête envoyée ! Vérifiez vos emails et votre dossier 'Etiquettes_Pelichet' dans quelques instants.");
        } catch (err) {
            alert("Erreur réseau");
            btn.innerText = "GÉNÉRER PDF";
        }
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
};
