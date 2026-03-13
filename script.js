const GAS_URL = "VOTRE_URL_ICI";

const fileInput = document.getElementById('excelFile');
const fileStatus = document.getElementById('fileStatus');

fileInput.onchange = () => { if(fileInput.files[0]) fileStatus.innerText = "Fichier : " + fileInput.files[0].name; };

document.getElementById('labelForm').onsubmit = async (e) => {
    e.preventDefault();
    const btn = document.getElementById('btnSubmit');
    const ldr = document.getElementById('loader');
    
    btn.disabled = true; ldr.classList.remove('hidden');

    const reader = new FileReader();
    reader.onload = async (event) => {
        try {
            const wb = XLSX.read(new Uint8Array(event.target.result), { type: 'array', cellText: true });
            const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { raw: false });
            const codes = json.map(r => r.Code || r.code || r.CODE).filter(c => c);

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

            // Appel GAS
            const response = await fetch(GAS_URL, { method: 'POST', mode: 'no-cors', body: JSON.stringify(payload) });

            btn.disabled = false; ldr.classList.add('hidden');
            document.getElementById('result').classList.remove('hidden');
            document.getElementById('driveLink').href = "https://drive.google.com/";
            alert("PDF envoyé sur Drive (et par email si renseigné)");

        } catch (err) {
            alert("Erreur : " + err.message);
            btn.disabled = false; ldr.classList.add('hidden');
        }
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
};
