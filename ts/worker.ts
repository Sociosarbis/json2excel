/// <reference lib="webworker" />
import "fast-text-encoding/text";
import wasmInit, { import_to_xlsx, init_panic_hook } from "../wasm/xlsx_import";
import { import_to_xlsx as import_to_xlsx_ts } from '../tsImpl';

type TableJson = {
    data: {
      name: string
      cells: ({ v?: string; s?: number } | null)[][]
      plain?: string[][]
      cols: { width: number }[]
      rows: { height: number }[]
      default_row_height?: number
      merged?: { from: { column: number; row: number }; to: { column: number; row: number } }[]
    }[]
    styles: Record<string, string>[]
}

type Config = {
    uid?: number,
    wasmPath?: string
    data: TableJson
}

declare var __buildVersion: string;

self.onmessage = function(e: MessageEvent) {
    if (e.data.type === "convert"){
        doConvert(e.data);
    }
}


let isLoaded = false;
function onData(result: ArrayBuffer | Uint8Array, uid?: number) {
    const blob = new Blob([result], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,"
    });

    self.postMessage({
        uid: uid || (new Date()).valueOf(),
        type: "ready",
        blob
    });
}
async function doConvert(config: Config){
    try {
        if ('WebAssembly' in self) {
            if (isLoaded) {
                const result = import_to_xlsx(config.data);
                onData(result, config.uid)
            } else {
                const path = `${config.wasmPath}?${__buildVersion}`;
                await wasmInit(path);
                isLoaded = true;
                init_panic_hook();
                doConvert(config);
            }
        } else {
            const result = await import_to_xlsx_ts(config.data);
            onData(result, config.uid); 
        }
    } catch (e) {
        self.postMessage({
            uid: config.uid || (new Date()).valueOf(),
            type: 'error',
            message: e.message
        })
    }
}