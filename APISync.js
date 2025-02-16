// Função principal que sincroniza os dados da planilha com o board do Monday.com
function syncSheetsForMonday() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getRange("A2:F" + sheet.getLastRow()).getValues(); // Obtém todas as linhas com dados

    const apiKey = "APIKEY"; // Defina sua chave de API
    const boardId = "BOARDIDMONDAY"; // ID do board no Monday.com
    const url = "https://api.monday.com/v2"; // URL da API do Monday.com

    try {
        // Obtém os itens existentes do board do Monday.com
        const existingItems = fetchExistingItems(apiKey, boardId, url);

        // Processa cada linha de dados da planilha
        processSpreadsheetData(data, existingItems, apiKey, boardId, url);
    } catch (e) {
        Logger.log("Erro durante a sincronização: " + e.toString());
    }
}

// Função que busca os itens existentes no board do Monday.com
function fetchExistingItems(apiKey, boardId, url) {
    const query = `{
        boards(ids: [${boardId}]) {
            id
            name
            items_page(limit: 500) {
                items {
                    id
                    name
                    column_values(ids: ["column_01", "column_02", "column_03", "column_04", "column_05"]) {
                        id
                        text
                    }
                }
            }
        }
    }`;

    const options = createRequestOptions(apiKey, query);
    const response = UrlFetchApp.fetch(url, options);
    const jsonResponse = JSON.parse(response.getContentText());
    Logger.log("Resposta da consulta dos itens: " + JSON.stringify(jsonResponse, null, 2));

    return mapExistingItems(jsonResponse);
}

// Função que cria as opções para a requisição HTTP
function createRequestOptions(apiKey, query) {
    return {
        method: 'post',
        headers: {
            'Authorization': 'Bearer ' + apiKey,
            'Content-Type': 'application/json'
        },
        payload: JSON.stringify({ query: query }),
        muteHttpExceptions: true
    };
}

// Função que mapeia os itens existentes e cria um objeto com as chaves
function mapExistingItems(jsonResponse) {
    const existingItems = {};
    const board = jsonResponse.data?.boards?.[0];

    if (board && board.items_page && Array.isArray(board.items_page.items)) {
        board.items_page.items.forEach(item => {
            const column02FromMonday = item.column_values.find(col => col.id === "column_02")?.text;
            if (column02FromMonday) {
                // Normaliza a chave (remove espaços e converte para minúsculo)
                existingItems[column02FromMonday.trim().toLowerCase()] = item.id;
            }
        });
    } else {
        Logger.log("Nenhum item encontrado no board.");
    }

    Logger.log("Itens existentes (map): " + JSON.stringify(existingItems, null, 2));
    return existingItems;
}

// Função que processa cada linha da planilha
function processSpreadsheetData(data, existingItems, apiKey, boardId, url) {
    data.forEach(row => {
        const column01 = String(row[0]).trim();
        const column02 = String(row[1]).trim().toLowerCase();
        const column03 = String(row[2]).trim();
        const column04 = String(row[3]).trim();
        const column05 = String(row[4]).trim();

        // Se o column_01 estiver vazio, preenche com "N/A"
        const column01ToSend = column01 === "" ? "N/A" : column01;

        Logger.log(`column_01: ${column01ToSend}, column_02: ${column02}, column_03: ${column03}, column_04: ${column04}, column_05: ${column05}`);

        // Prepara os valores das colunas para envio
        const columnValues = prepareColumnValues(column01ToSend, column02, column03, column04, column05);

        // Verifica se o item já existe, se sim, atualiza; caso contrário, cria um novo
        if (existingItems[column02]) {
            updateItem(existingItems[column02], boardId, columnValues, apiKey, url);
        } else {
            createItem(boardId, column01ToSend, columnValues, apiKey, url);
        }
    });
}

// Função que prepara os valores das colunas para o Monday.com
function prepareColumnValues(column01, column02, column03, column04, column05) {
    // Se algum valor for vazio, preenche com " " (em branco)
    const columnValuesObj = {
        "column_01": column01,  // Se column_01 for vazio, será "N/A"
        "column_02": column02,
        "column_03": column03,
        "column_04": column04 || " ",  // Preenche com " " se estiver vazio
        "column_05": column05 || " ",  // Preenche com " " se estiver vazio
    };

    // Retorna os valores formatados para JSON
    return JSON.stringify(columnValuesObj).replace(/"/g, '\\"');
}

// Função que atualiza um item existente no board do Monday.com
function updateItem(itemId, boardId, columnValues, apiKey, url) {
    const updateQuery = `mutation {
        change_multiple_column_values(item_id: ${itemId}, board_id: ${boardId}, column_values: "${columnValues}") {
            id
        }
    }`;
    const options = createRequestOptions(apiKey, updateQuery);

    const updateResponse = UrlFetchApp.fetch(url, options);
    Logger.log("Atualizado: " + itemId + " | Resposta: " + updateResponse.getContentText());
}

// Função que cria um novo item no board do Monday.com
function createItem(boardId, column01, columnValues, apiKey, url) {
    const createQuery = `mutation {
        create_item(board_id: ${boardId}, item_name: "${column01}", column_values: "${columnValues}") {
            id
        }
    }`;
    const options = createRequestOptions(apiKey, createQuery);

    const createResponse = UrlFetchApp.fetch(url, options);
    Logger.log("Criado: " + column01 + " | Resposta: " + createResponse.getContentText());
}
