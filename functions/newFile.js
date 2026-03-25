/** Inicia a execução lógica do script 
 * @param {"folder"|"file"|"docs"|"sheets"|"slides"|"forms"} params.type - Define o tipo de arquivo que será copiado
 * @param {string} [params.name] - Nome do novo item
 * @param {string} [params.from] - ID do arquivo de origem
 * @param {string} [params.to] - ID da pasta destino do novo item
 * @param {object} [params.replace] - Objeto Key-Value, onde a Key é o valor a se encontrar, e o Value é o valor a se substituir no arquivo  
 * @param {string} [params.format="url"] - Formato desejado do link retornado
 * @param {string} [params.savetosheet_fileid] - ID da planilha para onde se deseja salvar o link
 * @param {string} [params.savetosheet_tablename] - Nome da tabela da planilha para onde se deseja salvar o link
 * @param {string} [params.savetosheet_idcolname] - Nome da coluna de ID na planilha para onde se deseja salvar o link
 * @param {string} [params.savetosheet_findid] - ID da linha na planilha para onde se deseja salvar o link
 * @param {string} [params.savetosheet_findcol] - Nome da coluna onde ficará o link na planilha para onde se deseja salvar o link
 * @returns {object} Objeto com status HTTP, link (response) e observações
*/

function newFile(params = {}) {

    // Define resposta final do script
    let result = {
        status: 200,
        response: "Unhandled error",
        observation: undefined
    }

    let newItemID, newItemURL = undefined                           // Armazena ID e URL do novo arquivo ou pasta
    if (!params.to) params.to = DriveApp.getRootFolder().getId()    // Utiliza pasta raiz se sem destino

    // Cria ou copia com base no tipo
    try {
        switch (params.type) {

            // Caso pasta
            case "folder":
                newItemID = createFolder(params.name, params.to);  // Cria pasta
                break;

            case "file":
                newItemID = copyFile(params.name, params.from, params.to)   // Cria arquivo
                break;

            case "docs":
                newItemID = copyFile(params.name, params.from || "newDoc", params.to)       // Cria arquivo
                if (params.replace) replaceDocs(newItemID, params.replace)                  // Substitui valores caso solicitado
                break;

            case "sheets":
                newItemID = copyFile(params.name, params.from || "newSheet", params.to)     // Cria arquivo
                if (params.replace) replaceSheets(newItemID, params.replace)                // Substitui valores caso solicitado
                break;

            case "slides":
                newItemID = copyFile(params.name, params.from || "newSlides", params.to)    // Cria arquivo
                if (params.replace) replaceSlides(newItemID, params.replace)                // Substitui valores caso solicitado
                break;

            case "forms":
                newItemID = copyFile(params.name, params.from || "newForms", params.to)     // Cria arquivo
                if (params.replace) replaceForms(newItemID, params.replace)                 // Substitui valores caso solicitado
                break;

            default:
                throw Error(`Parameter "type" is not valid`)
        }run


        // Cria o link de retorno
        newItemURL = generateLink(newItemID, params.format)

    } catch (e) {

        // Caso erro, parar imediatamente
        result.status = 500
        result.response = e.message
        return result

    }

    // Caso seja solicitado salvar o link em uma planilha
    try {
        if (params.savetosheet_fileid) {

            // Carrega a planilha na memória
            let sheet = new Codex(
                params.savetosheet_fileid,
                params.savetosheet_tablename,
                params.savetosheet_idcolname,
                { mode: "minimal", columns: [params.savetosheet_findcol] }
            )

            // Armazena o URL da planilha
            sheet.get(params.savetosheet_findid)[params.savetosheet_findcol] = newItemURL

            // Salva a planilha
            sheet.commit()
        }
    } catch (e) {
        result.status = 201
        result.observation = `Unable to save to sheet: ${e.message}`
    }

    // Encerra a execução
    result.response = newItemURL
    return result
}


