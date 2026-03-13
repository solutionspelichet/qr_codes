const GAS_URL = "https://script.google.com/macros/s/AKfycbxDyfl3NhpEbFEL7zCdQgRCQrIutFqRwSnd1gGUOosaWIiIx5PdNQJbZPuSu_7mneUq/exec";

document.getElementById('labelForm').onsubmit = async (e) => {
    e.preventDefault();
    const btn = document.getElementById('btnSubmit');
    const ldr = document.getElementById('loader');
    const fileInput = document.getElementById('excelFile');

    if (!fileInput.files[0]) return alert("Fichier manquant");

    btn.disabled = true; ldr.classList.remove('hidden');

    const reader = new FileReader();
    reader.onload = async (event) => {
        try {
            const data = new Uint8Array(event.target.result);
            const wb = XLSX.read(data, { type: 'array', cellText: true, cellDates: true });
            const ws = wb.Sheets[wb.SheetNames[0]];
            
            // raw: false est vital pour garder le formatage visuel d'Excel
            const json = XLSX.utils.sheet_to_json(ws, { raw: false });

            const codes = json.map(r => r.Code || r.code || r.CODE).filter(c => c);

            const payload = {
                rows: codes,
                options: {
                    type: document.getElementById('type').value,
                    preset: document.getElementById('preset').value,
                    showText: document.getElementById('showText').value === "true",
                    textPos: document.getElementById('textPos').value
                }
            };

            await fetch(GAS_URL, {
                method: 'POST',
                mode: 'no-cors',
                body: JSON.stringify(payload)
            });

            btn.disabled = false; ldr.classList.add('hidden');
            document.getElementById('result').classList.remove('hidden');
        } catch (err) {
            alert("Erreur lecture: " + err.message);
            btn.disabled = false; ldr.classList.add('hidden');
        }
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
};
