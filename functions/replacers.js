

/** Substitui dados em um Documento Google
 * @param {string} id - ID do arquivo
 * @param {object} replace - Object Key-Value para substituição
 */
function replaceDocs(id, replace) {

    /** Substitui dados em uma aba.
     * @param {GoogleAppsScript.Document.DocumentTab} tab - Aba a se substituir
     */
    function replaceInTab(tab) {
        for (let search of Object.keys(replace)) {                                                          // Para cada valor:
            if (tab.getHeader()) tab.getHeader().replaceText(RE2escape(`{{${search}}}`), replace[search])   // Substitui no header
            if (tab.getBody()) tab.getBody().replaceText(RE2escape(`{{${search}}}`), replace[search])       // Substitui no body
            if (tab.getFooter()) tab.getFooter().replaceText(RE2escape(`{{${search}}}`), replace[search])   // Substitui no footer
        }
    }

    /** Substitui recursivamente em todas as abas.
     * @param {GoogleAppsScript.Document.Tab[]} tabs - Array com abas a se substituir
     */
    function recursiveReplace(tabs) {
        for (let tab of tabs) {                                         // Para cada aba:
            if (tab.getType() === DocumentApp.TabType.DOCUMENT_TAB) {   // Se aba de documento
                replaceInTab(tab.asDocumentTab())                       // Substitui dados na aba atual
            }
            recursiveReplace(tab.getChildTabs())    // Chama a si mesma nas abas-filha
        }
    }

    let doc = DocumentApp.openById(id)  // Abre o documento
    recursiveReplace(doc.getTabs())     // Substitui dados em todas as abas

}



/** Substitui dados em uma Planilha Google
 * @param {string} id - ID do arquivo
 * @param {object} replace - Object Key-Value para substituição
 */
function replaceSheets(id, replace) {

    let sheet = SpreadsheetApp.openById(id)         // Abre a planilha
    for (let search of Object.keys(replace)) {      // Para cada valor:
        sheet.createTextFinder(`{{${search}}}`)     // Encontre valor
            .ignoreDiacritics(false)                // Não ignorar diacríticos
            .matchCase(true)                        // Usar Case-sensitive
            .matchEntireCell(false)                 // Não validar dentro da célula toda
            .matchFormulaText(false)                // Não pesquisar dentro de fórmula
            .replaceAllWith(replace[search])        // Substituir todos os valores
    }
}

/** Substitui dados em um Slide Google
 * @param {string} id - ID do arquivo
 * @param {object} replace - Object Key-Value para substituição
 */
function replaceSlides(id, replace) {

    let slide = SlidesApp.openById(id)          // Abre o slide
    for (let search of Object.keys(replace)) {  // Para cada valor:
        slide.replaceAllText(                   // Substitui todos    
            `{{${search}}}`,                    // Encontre valor
            replace[search],                    // Substituir por este
            true                                // Usar Case-sensitive
        )
    }
}

/** Substitui dados em um Forms Google
 * @param {string} id - ID do arquivo
 * @param {object} replace - Object Key-Value para substituição
 */
function replaceForms(id, replace) {

    /** Substitui em uma string.
     * @param {string} originalText - String original
     * @returns {string} - String substituída
    */
    function replaceInString(originalText) {
        if (!originalText) return originalText                              // Ignora textos vazios
        let newText = originalText                                          // Cria var com novo texto
        for (let search of Object.keys(replace)) {                          // Para cada valor:
            newText = newText.replaceAll(`{{${search}}}`, replace[search]); // Aplicar substituição
        }
        return newText;
    }

    let form = FormApp.openById(id) // Abre formulário
    let oldValue, newValue          // Define vars de substituição

    // Edita título
    oldValue = form.getTitle()
    newValue = replaceInString(oldValue)
    if (oldValue !== newValue) form.setTitle(newValue)

    // Edita descrição
    oldValue = form.getDescription()
    newValue = replaceInString(oldValue)
    if (oldValue !== newValue) form.setDescription(newValue)

    // Edita todos os itens
    for (let item of form.getItems()) {

        // Título
        oldValue = item.getTitle()
        newValue = replaceInString(oldValue)
        if (oldValue !== newValue) item.setTitle(newValue)

        // Descrição
        oldValue = item.getHelpText()
        newValue = replaceInString(oldValue)
        if (oldValue !== newValue) item.setHelpText(newValue)
    }
}


