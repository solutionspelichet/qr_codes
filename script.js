const GAS_URL = "https://script.google.com/macros/s/AKfycbxDyfl3NhpEbFEL7zCdQgRCQrIutFqRwSnd1gGUOosaWIiIx5PdNQJbZPuSu_7mneUq/exec";

document.getElementById('excelFile').onchange = (e) => {
    if (e.target.files[0]) document.getElementById('fileStatus').innerText = "Fichier : " + e.target.files[0].name;
};

document.getElementById('labelForm').onsubmit = async (e) => {
    e.preventDefault();
    const btn = document.getElementById('btnSubmit');
    const ldr = document.getElementById('loader');
    
    btn.disabled = true; ldr.classList.remove('hidden');

    const reader = new FileReader();
    reader.onload = async (event) => {
        const data = new Uint8Array(event.target.result);
        const wb = XLSX.read(data, { type: 'array', cellText: true, cellDates: true }); // cellText force le formatage texte
        const ws = wb.Sheets[wb.SheetNames[0]];
        
        // Force Excel à ne pas convertir en nombres (évite le 1.0092E14)
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

        await fetch(GAS_URL, { method: 'POST', mode: 'no-cors', body: JSON.stringify(payload) });

        btn.disabled = false; ldr.classList.add('hidden');
        document.getElementById('result').classList.remove('hidden');
    };
    reader.readAsArrayBuffer(document.getElementById('excelFile').files[0]);
};
