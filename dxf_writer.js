// DXF檔案寫入器 - 用於匯出梁編號到AutoCAD
class DXFWriter {
    constructor() {
        this.lines = [];
        this.entities = [];
        this.layers = new Map();
        this.textStyles = new Map();
    }

    // 開始DXF檔案
    startDXF() {
        this.lines.push('0');
        this.lines.push('SECTION');
        this.lines.push('2');
        this.lines.push('HEADER');
        this.lines.push('9');
        this.lines.push('$ACADVER');
        this.lines.push('1');
        this.lines.push('AC1015');
        this.lines.push('0');
        this.lines.push('ENDSEC');
    }

    // 定義圖層
    addLayer(name, color = 7) {
        this.layers.set(name, color);
        this.lines.push('0');
        this.lines.push('LAYER');
        this.lines.push('2');
        this.lines.push(name);
        this.lines.push('70');
        this.lines.push('0');
        this.lines.push('62');
        this.lines.push(color.toString());
        this.lines.push('6');
        this.lines.push('CONTINUOUS');
    }

    // 定義文字樣式
    addTextStyle(name, fontName = 'Arial') {
        this.textStyles.set(name, fontName);
        this.lines.push('0');
        this.lines.push('STYLE');
        this.lines.push('2');
        this.lines.push(name);
        this.lines.push('70');
        this.lines.push('0');
        this.lines.push('40');
        this.lines.push('0.0');
        this.lines.push('41');
        this.lines.push('1.0');
        this.lines.push('50');
        this.lines.push('0.0');
        this.lines.push('71');
        this.lines.push('0');
        this.lines.push('42');
        this.lines.push('2.5');
        this.lines.push('3');
        this.lines.push(fontName);
    }

    // 添加文字實體
    addText(x, y, z, text, height = 2.5, layer = '0', style = 'STANDARD', rotation = 0) {
        this.lines.push('0');
        this.lines.push('TEXT');
        this.lines.push('8');
        this.lines.push(layer);
        this.lines.push('10');
        this.lines.push(x.toString());
        this.lines.push('20');
        this.lines.push(y.toString());
        this.lines.push('30');
        this.lines.push(z.toString());
        this.lines.push('40');
        this.lines.push(height.toString());
        this.lines.push('1');
        this.lines.push(text);
        this.lines.push('7');
        this.lines.push(style);
        this.lines.push('50');
        this.lines.push(rotation.toString());
    }

    // 添加多行文字
    addMText(x, y, z, text, height = 2.5, layer = '0', width = 0) {
        this.lines.push('0');
        this.lines.push('MTEXT');
        this.lines.push('8');
        this.lines.push(layer);
        this.lines.push('10');
        this.lines.push(x.toString());
        this.lines.push('20');
        this.lines.push(y.toString());
        this.lines.push('30');
        this.lines.push(z.toString());
        this.lines.push('40');
        this.lines.push(height.toString());
        this.lines.push('41');
        this.lines.push(width.toString());
        this.lines.push('1');
        this.lines.push(text);
    }

    // 添加線條
    addLine(x1, y1, z1, x2, y2, z2, layer = '0') {
        this.lines.push('0');
        this.lines.push('LINE');
        this.lines.push('8');
        this.lines.push(layer);
        this.lines.push('10');
        this.lines.push(x1.toString());
        this.lines.push('20');
        this.lines.push(y1.toString());
        this.lines.push('30');
        this.lines.push(z1.toString());
        this.lines.push('11');
        this.lines.push(x2.toString());
        this.lines.push('21');
        this.lines.push(y2.toString());
        this.lines.push('31');
        this.lines.push(z2.toString());
    }

    // 添加圓圈（用於標記梁位置）
    addCircle(x, y, z, radius, layer = '0') {
        this.lines.push('0');
        this.lines.push('CIRCLE');
        this.lines.push('8');
        this.lines.push(layer);
        this.lines.push('10');
        this.lines.push(x.toString());
        this.lines.push('20');
        this.lines.push(y.toString());
        this.lines.push('30');
        this.lines.push(z.toString());
        this.lines.push('40');
        this.lines.push(radius.toString());
    }

    // 結束DXF檔案
    endDXF() {
        this.lines.push('0');
        this.lines.push('ENDSEC');
        this.lines.push('0');
        this.lines.push('EOF');
    }

    // 生成DXF內容
    generateDXF() {
        this.startDXF();
        
        // 添加圖層定義
        this.lines.push('0');
        this.lines.push('SECTION');
        this.lines.push('2');
        this.lines.push('TABLES');
        
        // 添加圖層
        this.addLayer('BEAM_MAIN', 1);      // 紅色 - 大梁
        this.addLayer('BEAM_SECONDARY', 3); // 綠色 - 小梁
        this.addLayer('BEAM_SPECIAL', 6);   // 洋紅色 - 特殊梁
        this.addLayer('BEAM_LABELS', 7);    // 白色 - 標籤
        
        // 添加文字樣式
        this.addTextStyle('BEAM_TEXT', 'Arial');
        
        this.lines.push('0');
        this.lines.push('ENDSEC');
        
        // 添加實體
        this.lines.push('0');
        this.lines.push('SECTION');
        this.lines.push('2');
        this.lines.push('ENTITIES');
        
        return this.lines.join('\n');
    }

    // 匯出梁編號到DXF
    exportBeamLabels(beamData) {
        beamData.forEach(beam => {
            const x = beam.x || 0;
            const y = beam.y || 0;
            const z = beam.z || 0;
            const label = beam.label || '';
            const height = beam.height || 2.5;
            const rotation = beam.rotation || 0;
            
            // 根據梁類型選擇圖層
            let layer = 'BEAM_LABELS';
            if (label.startsWith('B') && !label.startsWith('b')) {
                layer = 'BEAM_MAIN';
            } else if (label.startsWith('b') || label.startsWith('fb')) {
                layer = 'BEAM_SECONDARY';
            } else if (label.startsWith('WB') || label.startsWith('FWB')) {
                layer = 'BEAM_SPECIAL';
            }
            
            // 添加標籤文字
            this.addText(x, y, z, label, height, layer, 'BEAM_TEXT', rotation);
            
            // 可選：添加位置標記圓圈
            if (beam.showMarker) {
                this.addCircle(x, y, z, 0.5, layer);
            }
        });
    }
}

// 使用範例
function exportToDXF(beamData) {
    const writer = new DXFWriter();
    writer.exportBeamLabels(beamData);
    const dxfContent = writer.generateDXF();
    
    // 下載DXF檔案
    const blob = new Blob([dxfContent], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'beam_labels.dxf';
    a.click();
    URL.revokeObjectURL(url);
}

// 匯出函數供HTML使用
window.exportToDXF = exportToDXF;
window.DXFWriter = DXFWriter;
