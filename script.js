const GAS_URL = "https://script.google.com/macros/s/AKfycbxDyfl3NhpEbFEL7zCdQgRCQrIutFqRwSnd1gGUOosaWIiIx5PdNQJbZPuSu_7mneUq/exec";

document.getElementById('labelForm').onsubmit = async (e) => {
    e.preventDefault();
    const btn = document.getElementById('btnText');
    const fileInput = document.getElementById('excelFile');
    
    if(!fileInput.files[0]) return alert("Fichier manquant !");
    
    btn.innerText = "GÉNÉRATION EN COURS...";
    btn.parentElement.classList.add('opacity-50', 'pointer-events-none');
    
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
                    showText: document.getElementById('textPos').value !== "none",
                    textPos: document.getElementById('textPos').value,
                    email: document.getElementById('userEmail').value,
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

            btn.innerText = "GÉNÉRER SUR GOOGLE DRIVE";
            btn.parentElement.classList.remove('opacity-50', 'pointer-events-none');
            document.getElementById('result').classList.remove('hidden');
        } catch (err) {
            alert("Erreur réseau : " + err.message);
            btn.innerText = "GÉNÉRER SUR GOOGLE DRIVE";
        }
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
};
