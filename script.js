const GAS_URL = "https://script.google.com/macros/s/AKfycbxDyfl3NhpEbFEL7zCdQgRCQrIutFqRwSnd1gGUOosaWIiIx5PdNQJbZPuSu_7mneUq/exec";

const fileInput = document.getElementById('excelFile');
const fileStatus = document.getElementById('fileStatus');
const labelForm = document.getElementById('labelForm');
const btnSubmit = document.getElementById('btnSubmit');
const loader = document.getElementById('loader');
const btnText = document.getElementById('btnText');
const resultDiv = document.getElementById('result');

fileInput.onchange = () => {
    if (fileInput.files.length > 0) {
        fileStatus.innerText = "Fichier : " + fileInput.files[0].name;
        fileStatus.classList.add('text-blue-600');
    }
};

labelForm.onsubmit = async (e) => {
    e.preventDefault();
    
    // Récupération sécurisée des éléments
    const type = document.getElementById('type').value;
    const preset = document.getElementById('preset').value;
    const showText = document.getElementById('showText').value === "true";
    const textPos = document.getElementById('textPos').value;

    if (!fileInput.files[0]) return alert("Veuillez choisir un fichier Excel");

    // UI State
    btnSubmit.disabled = true;
    loader.classList.remove('hidden');
    btnText.innerText = "TRAITEMENT...";
    resultDiv.classList.add('hidden');

    const reader = new FileReader();
    reader.onload = async (event) => {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(sheet);

            // On cherche la colonne 'Code' (insensible à la casse)
            const codes = json.map(r => r.Code || r.code || r.CODE).filter(c => c);

            if (codes.length === 0) throw new Error("Colonne 'Code' introuvable");

            const payload = {
                rows: codes,
                options: { type, preset, showText, textPos }
            };

            await fetch(GAS_URL, {
                method: 'POST',
                mode: 'no-cors',
                body: JSON.stringify(payload)
            });

            // Succès
            btnSubmit.disabled = false;
            loader.classList.add('hidden');
            btnText.innerText = "GÉNÉRER SUR GOOGLE DRIVE";
            resultDiv.classList.remove('hidden');

        } catch (err) {
            alert("Erreur : " + err.message);
            btnSubmit.disabled = false;
            loader.classList.add('hidden');
        }
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
};
