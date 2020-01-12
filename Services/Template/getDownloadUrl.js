var request = require('request');

module.exports = function getDownloadUrl(context, token, driveId, fileName) {
    context.log("TOKEN ", token);
    return new Promise((resolve, reject) => {
        context.log("FILENAME ", fileName);
        const url = `https://graph.microsoft.com/v1.0/drives/` + 
                    `${driveId}/root/children?$filter=name eq '${fileName}'`;
        context.log("DOWNLOAD URL ", url);
        try {
            
            request.get(url, {
                'auth': {
                    'bearer': token
                }
            }, (error, response, body) => {

                if (!error && response && response.statusCode == 200) {

                    const result = JSON.parse(response.body);
                    context.log("RESULT VALUE ",result.value);
                    context.log("RESULT VALUE[0] ",result.value[0]);
                    if (result.value && result.value[0]) {
                        //context.log("RESULT VALUE ",result.value);
                        //context.log("RESULT VALUE[0] ",result.value[0]);
                        resolve(result.value[0]["@microsoft.graph.downloadUrl"]);
                    } else {
                        reject(`File not found in getDownloadUrl: ${fileName}`);
                    }

                } else {

                    if (error) {
                        reject(`Error in getDownloadUrl: ${error}`);
                    } else {
                        let b = JSON.parse(response.body);
                        reject(`Error ${b.error.code} in getDownloadUrl: ${b.error.message}`);
                    }
                    
                }
            });
        } catch (ex) {
            reject(`Error in getDownloadUrl: ${ex}`);
        }


    });
}