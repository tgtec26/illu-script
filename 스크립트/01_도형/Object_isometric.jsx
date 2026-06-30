#target illustrator

function drawIsometricBox() {
    if (app.documents.length === 0) {
        alert("문서를 먼저 열어주세요.");
        return;
    }

    var doc = app.activeDocument;
    var sizeStepMm = 0.05;
    var minSliderMm = sizeStepMm;
    var maxSliderMm = 10;
    var previewGroup = null;

    // UI 입력창 생성
    var win = new Window("dialog", "등각 투상도 (Isometric) 상자 생성");
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];

    // 입력 패널
    var panel = win.add("panel", undefined, "크기 입력 (mm)");
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.margins = 20;

    var widthControl = addSizeControl(panel, "가로 (Width):", "5");
    var depthControl = addSizeControl(panel, "세로 (Depth):", "5");
    var heightControl = addSizeControl(panel, "높이 (Height):", "5");
    var inputW = widthControl.input;
    var inputD = depthControl.input;
    var inputH = heightControl.input;

    var previewCheck = win.add("checkbox", undefined, "미리보기");
    previewCheck.value = true;

    // 버튼 그룹
    var groupBtn = win.add("group");
    groupBtn.alignment = "center";
    var btnOk = groupBtn.add("button", undefined, "그리기", {name: "ok"});
    var btnCancel = groupBtn.add("button", undefined, "취소", {name: "cancel"});

    var result = null;

    btnOk.onClick = function() {
        var size = readSizeInputs(true);
        if (size === null) {
            return;
        }

        result = size;
        win.close(1);
    };

    btnCancel.onClick = function() {
        win.close(0);
    };

    previewCheck.onClick = updatePreview;
    updatePreview();

    var dialogResult = win.show();
    clearPreview();
    if (dialogResult === 1 && result !== null) {
        createIsometricBox(doc, result.w, result.d, result.h, true);
    }

    function addSizeControl(parent, label, defaultValue) {
        var control = {};
        var row = parent.add("group");
        row.orientation = "row";
        row.alignChildren = ["left", "center"];
        row.add("statictext", undefined, label);
        control.input = row.add("edittext", undefined, defaultValue);
        control.input.characters = 6;

        control.scrollbar = parent.add(
            "scrollbar",
            undefined,
            valueToStep(parseSize(defaultValue)),
            valueToStep(minSliderMm),
            valueToStep(maxSliderMm)
        );
        control.scrollbar.preferredSize.width = 300;
        control.scrollbar.stepdelta = 1;
        control.scrollbar.jumpdelta = 10;
        control.isSyncing = false;

        control.input.onChange = function() {
            syncScrollbar(control, parseSize(control.input.text));
            updatePreview();
        };
        control.input.onChanging = function() {
            updatePreview();
        };

        control.scrollbar.onChanging = function() {
            if (control.isSyncing) {
                return;
            }

            control.input.text = formatSize(stepToValue(control.scrollbar.value));
            updatePreview();
        };

        return control;
    }

    function parseSize(text) {
        return parseFloat(String(text).replace(",", "."));
    }

    function formatSize(value) {
        value = Math.round(value / sizeStepMm) * sizeStepMm;
        value = Math.max(sizeStepMm, value);
        return value.toFixed(2);
    }

    function valueToStep(value) {
        return Math.round(value / sizeStepMm);
    }

    function stepToValue(step) {
        return step * sizeStepMm;
    }

    function syncScrollbar(control, value) {
        if (control.isSyncing || isNaN(value) || value <= 0) {
            return;
        }

        var step = valueToStep(value);
        step = Math.max(valueToStep(minSliderMm), Math.min(valueToStep(maxSliderMm), step));
        control.isSyncing = true;
        control.scrollbar.value = step;
        control.isSyncing = false;
    }

    function readSizeInputs(showAlert) {
        var w = parseSize(inputW.text);
        var d = parseSize(inputD.text);
        var h = parseSize(inputH.text);

        if (isNaN(w) || isNaN(d) || isNaN(h) || w <= 0 || d <= 0 || h <= 0) {
            if (showAlert) {
                alert("올바른 양의 숫자를 입력해주세요.");
            }
            return null;
        }

        return {w: w, d: d, h: h};
    }

    function updatePreview() {
        clearPreview();
        if (!previewCheck.value) {
            return;
        }

        var size = readSizeInputs(false);
        if (size === null) {
            return;
        }

        previewGroup = createIsometricBox(doc, size.w, size.d, size.h, false);
        app.redraw();
    }

    function clearPreview() {
        if (previewGroup === null) {
            return;
        }

        try {
            previewGroup.remove();
        } catch (e) {}
        previewGroup = null;
    }
}

function createIsometricBox(doc, w_mm, d_mm, h_mm, selectResult) {
    // mm 단위를 pt 단위로 변환
    var mm2pt = 2.834645;
    var W = w_mm * mm2pt;
    var D = d_mm * mm2pt;
    var H = h_mm * mm2pt;

    // 등각 투상도 30도 각도 계산 (라디안)
    var cos30 = Math.cos(30 * Math.PI / 180);
    var sin30 = Math.sin(30 * Math.PI / 180);

    // 현재 보이는 화면(view)의 중앙 좌표 계산
    var centerPoint = doc.views[0].centerPoint;
    var cx = centerPoint[0];
    var cy = centerPoint[1];

    // 도형이 화면 중앙에 오도록 시작점(맨 아래 꼭짓점) 조정
    var startX = cx;
    var startY = cy - (H / 2);

    // 7개의 주요 꼭짓점 좌표 계산
    var pBottom = [startX, startY];
    var pRight = [startX + W * cos30, startY + W * sin30];
    var pLeft = [startX - D * cos30, startY + D * sin30];
    var pCenterTop = [startX, startY + H];
    var pTopRight = [startX + W * cos30, startY + H + W * sin30];
    var pTopLeft = [startX - D * cos30, startY + H + D * sin30];
    var pTopBack = [startX + (W - D) * cos30, startY + H + (W + D) * sin30];

    // 그룹 생성
    var group = doc.groupItems.add();
    group.name = "Isometric Box (" + w_mm + "x" + d_mm + "x" + h_mm + "mm)";

    var rightFace = makeFace(doc, group, [pBottom, pRight, pTopRight, pCenterTop]);
    var leftFace = makeFace(doc, group, [pBottom, pCenterTop, pTopLeft, pLeft]);
    var topFace = makeFace(doc, group, [pCenterTop, pTopRight, pTopBack, pTopLeft]);

    // 선과 면 스타일 지정
    var colorBlack = makeColor(doc, 0);
    var colorWhite = makeColor(doc, 100);

    var faces = [rightFace, leftFace, topFace];
    for (var i = 0; i < faces.length; i++) {
        faces[i].filled = true;
        faces[i].fillColor = colorWhite;
        faces[i].stroked = true;
        faces[i].strokeColor = colorBlack;
        faces[i].strokeWidth = 0.3; 
        faces[i].strokeJoin = StrokeJoin.ROUNDENDJOIN; // 모서리가 깔끔하게 맞물리도록 둥근 조인 사용
    }

    if (selectResult) {
        doc.selection = null;
        group.selected = true;
    }

    return group;
}

function makeFace(doc, group, points) {
    var face = doc.pathItems.add();
    face.setEntirePath(points);
    face.closed = true;
    face.move(group, ElementPlacement.PLACEATEND);
    return face;
}

function makeColor(doc, whitePercent) {
    var color;
    if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
        color = new CMYKColor();
        color.cyan = 0;
        color.magenta = 0;
        color.yellow = 0;
        color.black = 100 - whitePercent;
    } else {
        color = new RGBColor();
        color.red = Math.round(255 * whitePercent / 100);
        color.green = Math.round(255 * whitePercent / 100);
        color.blue = Math.round(255 * whitePercent / 100);
    }

    return color;
}

drawIsometricBox();
