import "../node_modules/fast-text-encoding/text";
import wasmInit, { import_to_xlsx } from "../wasm/xlsx_import";

onmessage = function(e) {
    if (e.data.type === "convert"){
        doConvert(e.data);
    }
}


let isLoaded = false;
function doConvert(config){
    if (isLoaded) {
        const result = import_to_xlsx(config.data);
        const blob = new Blob([result], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,"
        });

        postMessage({
            uid: config.uid || (new Date()).valueOf(),
            type: "ready",
            blob
        });
    } else {
        const path = config.wasmPath || "https://cdn.dhtmlx.com/libs/json2excel/1.0/lib.wasm";

        wasmInit(path).then(() => {
            isLoaded = true;
            doConvert(config);
        }).catch(e => console.log(e));
    }
}