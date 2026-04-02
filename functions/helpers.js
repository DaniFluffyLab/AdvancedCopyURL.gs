
/** Cria nova pasta.
 * @param {string} name - Nome da nova pasta
 * @param {string} destID - ID da pasta destino
*/
function createFolder(name, destID) {
    if (!name) name = `New folder ${(new Date()).toLocaleString()}` // Cria nome genérico se sem nome
    let destFolder = DriveApp.getFolderById(destID)                 // Obtém pasta de destino
    let newFolder = destFolder.createFolder(name)                   // Cria nova pasta
    return newFolder.getId()                                        // Retorna ID da nova pasta
}



/** Duplica arquivo.
 * @param {string} name - Nome do novo arquivo
 * @param {string} originID - ID do arquivo original
 * @param {string} destID - ID da pasta destino
*/
function copyFile(name, originID, destID) {

    // Obtém pasta de destino
    let destFolder = DriveApp.getFolderById(destID)

    // Cria arquivo novo
    let newFile = undefined;
    switch (originID) {
        case "newDoc": newFile = DocumentApp.create("Untitled document").getId(); break;
        case "newSheet": newFile = SpreadsheetApp.create("Untitled sheet").getId(); break;
        case "newSlides": newFile = SlidesApp.create("Untitled slide").getId(); break;
        case "newForms": newFile = FormApp.create("Untitled form").getId(); break;
        default: newFile = DriveApp.getFileById(originID).makeCopy().getId();
    }
    if (name) DriveApp.getFileById(newFile).setName(name)   // Renomeia arquivo se informado nome
    DriveApp.getFileById(newFile).moveTo(destFolder)        // Move para a pasta solicitada
    return newFile                                          // Retorna ID do novo arquivo
}




/**
 * Escapa uma string para ser usada como correspondência exata 
 * em APIs do Google que utilizam o motor de Regex RE2.
 * * @param {string} text - O texto original (ex: "[Dados] planilha.docx")
 * @return {string} O texto escapado para RE2 (ex: "\[Dados\] planilha\.docx")
 */
function RE2escape(text) {
    if (!text) return "";
    return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Converte um objeto JavaScript em um novo objeto onde todos os valores são convertidos em strings.
 *
 * @param {object} obj O objeto JavaScript a ser convertido.
 * @returns {object} Um novo objeto com todos os valores convertidos para string.
 * @throws {TypeError} Se o argumento não for um objeto.
 */
function objValuesToString(obj) {

    /** Helper para converter objeto em string */
    function helper_convert(value, depth = 0) {

        // Vars usadas no switch
        let convertedValue, jsonValue;

        // Converte dado baseado no tipo
        switch (typeof value) {

            // Converte tipos simples
            case 'number':
            case 'boolean':
            case 'bigint':
            case 'string':
                return String(value);

            // Converte undefined em string vazia
            case 'undefined':
                return "";

            // Para tipos avançados
            case 'object':

                // NULL - String vazia
                if (value === null) return ""

                // DATAS - Data formatada
                if (value instanceof Date) return value.toLocaleString()

                // REGEX - Converte para string
                if (value instanceof RegExp) return value.toString()

                // ARRAY ou SET - Converte em JSON
                if ((value instanceof Set || value instanceof Array)) {
                    convertedValue = [...value]                     // Cria cópia de segurança
                    jsonValue = JSON.stringify(convertedValue)      // Converte dados
                    return jsonValue                                // Retorna resultados
                }

                // MAP - Converte em Objeto JSON
                if (value instanceof Map) {

                    convertedValue = [...value.entries()]                               // Obtém o encadeamento chave / valor
                    jsonValue = JSON.stringify(Object.fromEntries(convertedValue))      // Converte para JSON
                    return jsonValue                                                    // Retorna resultado
                }

                // OBJETO LITERAL - Converte em JSON
                if (Object.prototype.toString.call(value) === '[object Object]') return JSON.stringify(value)

            default:
                // Retorna string convertida de forma genérica
                return String(value)
        }
    }

    // Objeto de saída
    let convertedObj = {}

    // Converte todos os valores
    for (let key of Object.keys(obj)) convertedObj[key] = helper_convert(obj[key])
    
    // Retorna objeto
    return convertedObj
}

/**
 * Gera o link do arquivo no formato de exportação desejado.
 * @param {string} id - ID do arquivo
 * @param {string} [format="url"] - Formato desejado (ex: "url", "pdf", "docx", "xlsx", "download", "published")
 * @returns {string} Link formatado
 */
function generateLink(id, format = "url") {
    try {

        // Caso apenas URL, retorna valor diretamente
        if (format === "url") return `https://drive.google.com/open?id=${id}`;

        switch (DriveApp.getFileById(id).getMimeType()) {

            // Caso Google Docs
            case "application/vnd.google-apps.document":
                if (format === "docx") return `https://docs.google.com/document/d/${id}/export?format=docx`;
                if (format === "pdf") return `https://docs.google.com/document/d/${id}/export?format=pdf`;
                if (format === "odt") return `https://docs.google.com/document/d/${id}/export?format=odt`;
                if (format === "txt") return `https://docs.google.com/document/d/${id}/export?format=txt`;
                if (format === "rtf") return `https://docs.google.com/document/d/${id}/export?format=rtf`;
                if (format === "html") return `https://docs.google.com/document/d/${id}/export?format=html`;
                if (format === "epub") return `https://docs.google.com/document/d/${id}/export?format=epub`;
                if (format === "md") return `https://docs.google.com/document/d/${id}/export?format=md`;
                if (format === "download") return `https://docs.google.com/document/d/${id}/export?format=docx`;
                return `https://drive.google.com/open?id=${id}`;

            // Caso Google Planilhas
            case "application/vnd.google-apps.spreadsheet":
                if (format === "xlsx") return `https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx`;
                if (format === "ods") return `https://docs.google.com/spreadsheets/d/${id}/export?format=ods`;
                if (format === "pdf") return `https://docs.google.com/spreadsheets/d/${id}/export?format=pdf`;
                if (format === "html") return `https://docs.google.com/spreadsheets/d/${id}/export?format=html`;
                if (format === "csv") return `https://docs.google.com/spreadsheets/d/${id}/export?format=csv`;
                if (format === "tsv") return `https://docs.google.com/spreadsheets/d/${id}/export?format=tsv`;
                if (format === "download") return `https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx`;
                return `https://drive.google.com/open?id=${id}`;

            // Caso Google Apresentações
            case "application/vnd.google-apps.presentation":
                if (format === "pptx") return `https://docs.google.com/presentation/d/${id}/export/pptx`;
                if (format === "odp") return `https://docs.google.com/presentation/d/${id}/export/odp`;
                if (format === "pdf") return `https://docs.google.com/presentation/d/${id}/export/pdf`;
                if (format === "txt") return `https://docs.google.com/presentation/d/${id}/export/txt`;
                if (format === "jpg") return `https://docs.google.com/presentation/d/${id}/export/jpg`;
                if (format === "png") return `https://docs.google.com/presentation/d/${id}/export/png`;
                if (format === "svg") return `https://docs.google.com/presentation/d/${id}/export/svg`;
                if (format === "download") return `https://docs.google.com/presentation/d/${id}/export/pptx`;
                return `https://drive.google.com/open?id=${id}`;

            // Caso Google Desenhos
            case "application/vnd.google-apps.drawing":
                if (format === "pdf") return `https://docs.google.com/drawings/d/${id}/export/pdf`;
                if (format === "jpg") return `https://docs.google.com/drawings/d/${id}/export/jpg`;
                if (format === "png") return `https://docs.google.com/drawings/d/${id}/export/png`;
                if (format === "svg") return `https://docs.google.com/drawings/d/${id}/export/svg`;
                if (format === "download") return `https://docs.google.com/drawings/d/${id}/export/png`;
                return `https://drive.google.com/open?id=${id}`;

            // Caso Google Forms
            case "application/vnd.google-apps.form":
                if (format === "published") return FormApp.openById(id).getPublishedUrl();
                return `https://drive.google.com/open?id=${id}`;

            // Caso outros formatos
            default:
                if (format === "download") return DriveApp.getFileById(id).getDownloadUrl() || `https://drive.google.com/open?id=${id}`;

        }

    } catch (e) {

        // Fallback caso falha
        return `https://drive.google.com/open?id=${id}`;
    }
}