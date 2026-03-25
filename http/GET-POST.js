/**
 * Lida com requisições GET (Acessos pelo navegador ou chamadas de API)
 */
function doGet(e) {

    let params = e.parameter || {};

    // MODO API
    if (params.api === "true") {
        let result = newFile(params);
        return ContentService.createTextOutput(JSON.stringify(result))
            .setMimeType(ContentService.MimeType.JSON);
    }

    // WEB APP
    let html = HtmlService.createTemplateFromFile('http/index.html')
    html.params = params
    return html.evaluate()
        .setTitle(`Criando ${params.name || "arquivo"}...`)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}




/**
 * Lida com requisições POST (Webhooks e chamadas de sistemas externos)
 */
function doPost(e) {
    let params = {};

    try {

        // Interpreta corpo da requisição como parâmetro
        if (e.postData && e.postData.contents) {
            params = JSON.parse(e.postData.contents);
        }
        
        // Interpreta parâmetros de URL como parâmetro
        else {
            params = e.parameter;
        }

    } 
    
    // Retorna erro caso não ache parâmetros
    catch (err) {
        return ContentService.createTextOutput(JSON.stringify({
            status: 400,
            response: "Invalid JSON payload",
            observation: err.message
        })).setMimeType(ContentService.MimeType.JSON);
    }

    // MODO API
    let result = newFile(params);
    return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
}