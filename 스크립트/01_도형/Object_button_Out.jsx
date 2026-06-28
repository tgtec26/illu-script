// Button Projection Script for Adobe Illustrator
// 선택한 원/타원을 뒤쪽으로 복제하고 접선 라인을 추가해 버튼처럼 입체화

(function() {
    if (app.documents.length === 0) {
        alert("열려있는 문서가 없습니다.");
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (!sel || sel.length === 0 || sel[0].typename !== "PathItem") {
        alert("원 또는 타원 패스를 선택해주세요.");
        return;
    }

    var frontFace = sel[0];
    if (!frontFace.closed) {
        alert("닫힌 원 또는 타원 패스를 선택해주세요.");
        return;
    }

    var bounds = frontFace.geometricBounds; // [left, top, right, bottom]
    var width = bounds[2] - bounds[0];
    var height = bounds[1] - bounds[3];
    var mmToPt = 2.83464567;
    var defaultDepthMm = Math.round(Math.min(width, height) / 5 / mmToPt * 100) / 100;
    var previewItems = [];

    var depthMm = showDepthDialog(defaultDepthMm, function(valueMm) {
        clearPreview();
        previewItems = createButton(valueMm * mmToPt);
        app.redraw();
    }, clearPreview);

    clearPreview();
    if (depthMm === null) {
        return;
    }

    createButton(depthMm * mmToPt);
    doc.selection = null;

    function createButton(depth) {
        frontFace.strokeJoin = StrokeJoin.ROUNDENDJOIN;

        var backFace = frontFace.duplicate();
        backFace.name = "ButtonDepthOut";
        backFace.translate(depth, depth);
        backFace.strokeJoin = StrokeJoin.ROUNDENDJOIN;
        backFace.move(frontFace, ElementPlacement.PLACEAFTER);

        var tangentPoints = getEllipseTangentPoints(bounds, depth, depth);
        var tangentA = makeTangentLine(tangentPoints[0], depth, depth);
        var tangentB = makeTangentLine(tangentPoints[1], depth, depth);

        try {
            tangentA.move(frontFace, ElementPlacement.PLACEAFTER);
            tangentB.move(frontFace, ElementPlacement.PLACEAFTER);
        } catch (e) {}

        return [backFace, tangentA, tangentB];
    }

    function showDepthDialog(defaultValue, onPreview, onClearPreview) {
        var dialog = new Window("dialog", "버튼 깊이");
        dialog.orientation = "column";
        dialog.alignChildren = "fill";

        var inputGroup = dialog.add("group");
        inputGroup.add("statictext", undefined, "뒤로 이동 거리(mm)");
        var input = inputGroup.add("edittext", undefined, String(defaultValue));
        input.characters = 8;

        var previewCheck = dialog.add("checkbox", undefined, "미리보기");
        previewCheck.value = true;

        var buttons = dialog.add("group");
        buttons.alignment = "right";
        var okButton = buttons.add("button", undefined, "확인", {name: "ok"});
        var cancelButton = buttons.add("button", undefined, "취소", {name: "cancel"});

        var result = null;

        function readValue(showAlert) {
            var value = parseFloat(String(input.text).replace(",", "."));
            if (isNaN(value) || value <= 0) {
                if (showAlert) {
                    alert("0보다 큰 숫자를 입력해주세요.");
                }
                return null;
            }
            return value;
        }

        function updatePreview() {
            if (!previewCheck.value) {
                onClearPreview();
                return;
            }

            var value = readValue(false);
            if (value === null) {
                onClearPreview();
                return;
            }

            onPreview(value);
        }

        input.onChanging = updatePreview;
        previewCheck.onClick = updatePreview;
        okButton.onClick = function() {
            var value = readValue(true);
            if (value === null) {
                return;
            }
            result = value;
            dialog.close();
        };
        cancelButton.onClick = function() {
            result = null;
            dialog.close();
        };

        updatePreview();
        dialog.show();

        return result;
    }

    function clearPreview() {
        for (var i = previewItems.length - 1; i >= 0; i--) {
            try {
                previewItems[i].remove();
            } catch (e) {}
        }
        previewItems = [];
    }

    function getEllipseTangentPoints(itemBounds, dx, dy) {
        var left = itemBounds[0];
        var top = itemBounds[1];
        var right = itemBounds[2];
        var bottom = itemBounds[3];
        var cx = (left + right) / 2;
        var cy = (top + bottom) / 2;
        var rx = (right - left) / 2;
        var ry = (top - bottom) / 2;
        var angle = Math.atan2(-ry * dx, rx * dy);

        var p1 = [
            cx + (rx * Math.cos(angle)),
            cy + (ry * Math.sin(angle))
        ];
        var p2 = [
            cx + (rx * Math.cos(angle + Math.PI)),
            cy + (ry * Math.sin(angle + Math.PI))
        ];

        return [p1, p2];
    }

    function makeTangentLine(startPoint, dx, dy) {
        var line = doc.pathItems.add();
        line.setEntirePath([
            [startPoint[0], startPoint[1]],
            [startPoint[0] + dx, startPoint[1] + dy]
        ]);
        line.closed = false;
        line.filled = false;
        line.stroked = true;

        if (frontFace.stroked) {
            line.strokeColor = frontFace.strokeColor;
            line.strokeWidth = frontFace.strokeWidth;
            line.strokeDashes = frontFace.strokeDashes;
            line.strokeCap = frontFace.strokeCap;
            line.strokeJoin = StrokeJoin.ROUNDENDJOIN;
        }

        return line;
    }
})();
