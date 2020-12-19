const {GoogleSpreadsheet} = require('google-spreadsheet');
const docId = '1wGdeBw4NGt_Z8fSs6i4VZKDtmTsREiMNhPRcHKgRQrA';
const doc = new GoogleSpreadsheet(docId);

(async function() {    
    //utiliza as credenciais de contas de serviço
    doc.useServiceAccountAuth({
        "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCla0ZhUzNx+F57\nI7LnzMYPlyX8rfPne5UuQm2B7WENrX0ElfLadxXRiv9AI0w+CvdS/JltJXWXWqDW\nBfMnATfEkpXpt7tNwMZMxhxokZ/nVjOtDjD5OdnEhA05k0HdhAHHy8briHM87AOt\n/OiDye7Y6jhVIRHwCaeen+257ruhy4GFVzoIMP9+HUsONjdKO86eVQOmM9NJk1rB\niGVe364w4m961VoAVkpmSSU6BOwOVXiPIL+lPl6D1yIxERnsgw54snQktGw0dU2v\nafIBMvXPSOv9kOyhTEcBi9EGrLc8o/QzzHuMTW5tnK41HcnyQNo4PMm+KDFxxoe1\ntTQGBmTxAgMBAAECggEAJjPtt6lu7qnVLCw087rDnTMjp0UHWNUeQWY/Aywu52lD\nP61fvluVUuT+iTH9uDBxKR3EU0Q88Z0RGwyZuM7bsc7Gx4jSvaTBR1bTlcTYAKXU\nXmmyHsThCbUTltHu+rkzbhCeWqQTNSUuvl5z1ofq3PbO1r5B9cVNDGHUFcZQWA+Y\nNhWz8TWbToq1y5FIih0RcTfneqxZCGSRUOAqFrCMZNdDWe+k4+3ofv3Gt7RGSnNS\nz4CsQSvpBAYwXiMZnnyprIEPkX6Kfzt7WySskhj3pfYCkQ9tNp3+zmLMdqCr5oSx\nQc9PN6Xrb8jmRXj5JFCHcPQtNCgsN3c1NBhSo2cOLwKBgQDVr3FVi5hEskXKFiac\nsVCGnZJ3+0yFG0ToY44zyhCVDjl+6S4P57egrY1/E4bypp7DNVNAswyAeRjuR7KV\nJ6m8hFhvO9hRUaqkz52dB8aTJFdDn5JIaFIAoNnMdU2bLPUNx8VqnJbBgLUKb744\n9IiQ36simtPR7rxTRgRquqsnNwKBgQDGLQXZ0MaxX2pErL378miYYZ6vT/GRkfYX\n9HXeWb5PS6XmgqJPrS+s5QjH/s/wYLWc8ETT0GgSutR0brorrHIki97iiig54D8X\n4ua/byKG3VpleuRlHbv7flf4jUJY2uBiq9na2t+UjI7F9MMI8qnsG8dK5qcz1oCv\nBeB86nqZFwKBgQDIvbOLuMNILe2wJlUJuO28OMVDX9oH5ZE7e2M4tegUDzPmTKqQ\nGJACK2iU68RHqk3VdwAJ9OqWuqy4FLTouEUVq4LkpGTYKA9WGxCnV4mt62LrTToA\nObhnjLRvBfftAjQISRbly8s4Z3AsKMOb/+VXrDe6H5dETbGvzUQS++ATywKBgBhl\niobaENvdJzP1IB5YJVA9FE/4w4BsO6OPUMNiwO76HR5XjqvIYkoimAYm9GpfPXxo\nh9Cbo3RK08TRrNGblSGypmm1IGafmKTUJhwDDnkT3wEHM/7OvkmjsCjFGxndOCpt\nhZBPyZ57/0eXbjs3xHtwoAQ0iPj0uzrQumYmZ5lZAoGAdXDfVDmz/3Dc6kryxijD\nNZl1p5OkbcA9suDNNrTGB0wrPNs0WZMIQEqIGeswCCwzJJPp1BejbibopvhAbrF+\nbAwfSys8FNkYsmfunFH4o2Y1Mj0WTLtSeTyTmu1dcSgzLPwryHtThlXI0aSPoJbB\n83JYZMGHJu52DFyb26HS+lI=\n-----END PRIVATE KEY-----\n",
        "client_email": "feecaragua@wygor-299023.iam.gserviceaccount.com"
    })    

    //obtem as informações da planilha
    await doc.loadInfo();    
    const sheet = doc.sheetsByIndex[0];
    await sheet.loadCells('A2:H27');

    //constantes e variáveis necessárias para o processamento
    const totalAulas = sheet.getCellByA1('A2').value.match(/\d+/g).map(n => parseInt(n));
    let aluno    = undefined;
    let faltas   = undefined;
    let p1       = undefined;
    let p2       = undefined;
    let p3       = undefined;
    let situacao = undefined;
    let naf      = undefined;
    let aprovadoFalta = undefined;
    let media = undefined;

    //do the magic
    for (row = 4; row < 28; row++) {
        aluno    = sheet.getCellByA1('B' + row);
        faltas   = sheet.getCellByA1('C' + row);
        p1       = sheet.getCellByA1('D' + row);
        p2       = sheet.getCellByA1('E' + row);
        p3       = sheet.getCellByA1('F' + row);
        situacao = sheet.getCellByA1('G' + row);
        naf      = sheet.getCellByA1('H' + row);

        aprovadoFalta = (faltas.value * (totalAulas[0] / 100)) <= 25;
        media = Math.round((p1.value + p2.value + p3.value) / 3);

        if (!aprovadoFalta) {
            situacao.value = 'Reprovado por Falta';
        } else {
            if (media < 50) {
                situacao.value = 'Reprovado por Nota';
                naf.value      = 0;
            } else if (media >= 50 && media < 70) {
                situacao.value = 'Exame Final';
                naf.value      = 100 - media;       
            } else {
                situacao.value = 'Aprovado';
                naf.value      = '0';
            }  
        }
        situacao.horizontalAlignment = "LEFT";
        naf.horizontalAlignment = "RIGHT";

        //salva as alterações
        await sheet.saveUpdatedCells();   
        console.log('Aluno ', aluno.value, ' - Faltas: ', faltas.value, ' - Média: ', media, ' - Situação: ', situacao.value, ' - NAF: ', naf.value);  
    }
}());
