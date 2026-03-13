const GAS_URL = "https://script.google.com/macros/s/AKfycbxDyfl3NhpEbFEL7zCdQgRCQrIutFqRwSnd1gGUOosaWIiIx5PdNQJbZPuSu_7mneUq/exec";

document.getElementById('labelForm').onsubmit = async (e) => {
    e.preventDefault();
    const btn = document.getElementById('btnSubmit');
    const btnText = document.getElementById('btnText');
    const fileInput = document.getElementById('excelFile');
    const resultBox = document.getElementById('result');
    const errorBox = document.getElementById('errorBox');
    
    if(!fileInput.files[0]) return alert("Veuillez charger un fichier Excel.");
    
    btnText.innerText = "CRÉATION EN COURS (Patientez 10-20s)..."; 
    btn.disabled = true;
    resultBox.classList.add('hidden');
    errorBox.classList.add('hidden');

    const reader = new FileReader();
    reader.onload = async (event) => {
        try {
            const wb = XLSX.read(new Uint8Array(event.target.result), {type: 'array', cellText: true});
            const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {raw:false});
            const codes = data.map(r => r.Code || r.code || r.CODE).filter(c => c);

            if(codes.length === 0) throw new Error("Aucune donnée trouvée dans la colonne 'Code'.");

            const payload = {
                rows: codes,
                options: {
                    type: document.getElementById('type').value,
                    preset: document.getElementById('preset').value,
                    textPos: document.getElementById('textPos').value,
                    showText: document.getElementById('textPos').value !== "none",
                    email: document.getElementById('userEmail').value,
                    custW: document.getElementById('custW').value,
                    custH: document.getElementById('custH').value,
                    custCols: document.getElementById('custCols').value,
                    custRows: document.getElementById('custRows').value
                }
            };

            // On envoie en texte brut pour éviter les blocages CORS tout en pouvant lire la réponse JSON
            const response = await fetch(GAS_URL, { 
                method: 'POST', 
                body: JSON.stringify(payload)
                // ATTENTION: Nous avons retiré mode: 'no-cors'.
            });

            const result = await response.json();

            if (result.success) {
                resultBox.classList.remove('hidden');
                document.getElementById('driveLink').href = result.url;
                document.getElementById('emailStatus').innerText = "Email : " + result.emailInfo;
            } else {
                errorBox.classList.remove('hidden');
                document.getElementById('errorText').innerText = "Erreur Serveur : " + result.error;
            }

        } catch (err) {
            errorBox.classList.remove('hidden');
            document.getElementById('errorText').innerText = "Erreur Technique : " + err.message;
        } finally {
            btnText.innerText = "GÉNÉRER LE PDF"; 
            btn.disabled = false;
        }
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
};
