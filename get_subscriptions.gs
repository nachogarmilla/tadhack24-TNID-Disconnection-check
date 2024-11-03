function fetchB2cSubscriptions() {
  const url = 'https://api.staging.v2.tnid.com/company';
  const query = `
    query B2cSubscriptions {
      b2cSubscriptions {
        records {
          ... on B2cEmailSubscription {
            email
            expirationDate
            id
            insertedAt
            notificationEndTime
            notificationStartTime
            shareInformationWithTheCompany
            stopAllCommunications
            subscriptionTopics
            timezone
            type
            updatedAt
          }
          ... on B2cSmsSubscription {
            expirationDate
            id
            insertedAt
            notificationEndTime
            notificationStartTime
            shareInformationWithTheCompany
            stopAllCommunications
            subscriptionTopics
            telephoneNumber
            timezone
            type
            updatedAt
          }
          ... on B2cVoiceSubscription {
            expirationDate
            id
            insertedAt
            notificationEndTime
            notificationStartTime
            shareInformationWithTheCompany
            stopAllCommunications
            subscriptionTopics
            telephoneNumber
            timezone
            type
            updatedAt
          }
        }
      }
    }
  `;

  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify({ query: query }),
    headers: {
      'Authorization': 'Bearer ' + getToken() // Obtén el token de la hoja de cálculo
    }
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    // Procesar los datos y escribirlos en la hoja de cálculo
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const records = json.data.b2cSubscriptions.records;

    // Limpiar la hoja antes de insertar nuevos datos
    sheet.clear();
    // Escribir los encabezados
    sheet.appendRow(['ID', 'Contact', 'Inserted At','Updated At','Expiration Date', 'Stop All Communications', 'Notification Start Time', 'Notification End Time', 'Type', 'Telephone Number', 'Timezone', 'Subscription Topics']);

    // Escribir los datos
    records.forEach(record => {
      const row = [
        record.id,
        record.email || record.telephoneNumber, // Usa email o número de teléfono según el tipo
        record.insertedAt,
        record.updatedAt,
        record.expirationDate,
        record.stopAllCommunications,
        record.notificationStartTime,
        record.notificationEndTime,
        record.type,
        record.telephoneNumber || '', // Solo incluir si es un SMS o voz
        record.timezone,

        record.subscriptionTopics ? record.subscriptionTopics.join(', ') : '' // Convierte el array a una cadena
      ];
      sheet.appendRow(row);
    });

  } catch (error) {
    Logger.log('Error fetching data: ' + error);
  }
}

function getToken() {
  const tokenSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('token');
  const token = tokenSheet.getRange('B1').getValue();
  return token;
}
