const xlsx = require('xlsx');
const planilhas = [
    './amostras/005-Dataset-Amostra-A.xlsx',
    './amostras/005-Dataset-Amostra-B.xlsx',
    './amostras/005-Dataset-Amostra-C.xlsx',
    './amostras/005-Dataset-Amostra-D.xlsx',
];
var filePath = '';
var pontos = '';
var resultado = [];

class dados {
    constructor(nome, eqReta, r2, xM, yM) {
        this.nome = nome;
        this.eqReta = eqReta;
        this.r2 = r2;
        this.xM = xM;
        this.yM = yM;
    }
}

function calcularEquacaoDaReta(pontos) {
    let n = pontos.length;

    let somaX = 0;
    let somaY = 0;
    let somaXY = 0;
    let somaX2 = 0;

    for (let i = 0; i < n; i++) {
        let x = pontos[i][0];
        let y = pontos[i][1];
        somaX += x;
        somaY += y;
        somaXY += x * y;
        somaX2 += x * x;
    }

    let m = (n * somaXY - somaX * somaY) / (n * somaX2 - somaX * somaX);
    let b = (somaY - m * somaX) / n;

    // Calcular R²
    let somaResiduos = 0;  // Soma dos resíduos
    let somaTotal = 0;     // Soma total dos quadrados

    for (let i = 0; i < n; i++) {
        let x = pontos[i][0];
        let y = pontos[i][1];
        let yEstimado = m * x + b;
        somaResiduos += Math.pow(y - yEstimado, 2);
        somaTotal += Math.pow(y - (somaY / n), 2);
    }

    let r2 = 1 - (somaResiduos / somaTotal);
    let mediaX = somaX / n;
    let mediaY = somaY / n;

    return { m, b, r2, mediaX, mediaY };
}

// Ler os dados do arquivo Excel
function lerExcel(filePath) {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0]; // Primeiro sheet
    const worksheet = workbook.Sheets[sheetName];

    // Converter os dados da planilha em um array de pontos
    const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    const pontos = data.slice(1).map(row => {
        if (row.length < 2) return null; // Ignora linhas com menos de 2 colunas

        const x = row[0] !== undefined ? parseFloat(row[0].toString().replace(',', '.')) : null;
        const y = row[1] !== undefined ? parseFloat(row[1].toString().replace(',', '.')) : null;

        return (x !== null && y !== null) ? [x, y] : null; // Retorna null se x ou y não forem válidos
    }).filter(row => row !== null); // Remove linhas nulas

    return pontos;
}

for (let i=0; i<planilhas.length; i++) {
    filePath = planilhas[i];
    pontos = lerExcel(filePath);

    if (pontos.length < 2) {
        console.error("Não há dados suficientes para calcular a equação da reta.");
    } else {
        // Calcular a equação da reta
        let { m, b, r2, mediaX, mediaY } = calcularEquacaoDaReta(pontos);        
        dado = new dados (filePath, `y = ${m.toFixed(2)}x + ${b.toFixed(2)}`, parseFloat(r2.toFixed(4)), parseFloat(mediaX.toFixed(2)), parseFloat(mediaY.toFixed(2)));
        resultado[i] = dado;
    }
}

console.table(resultado);
