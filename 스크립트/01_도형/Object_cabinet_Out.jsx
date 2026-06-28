// Cabinet Projection Script for Adobe Illustrator
// 선택한 사각형을 캐비넷 투영법으로 입체화

(function() {
    if (app.documents.length === 0) {
        alert("열려있는 문서가 없습니다.");
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (!sel || sel.length === 0 || sel[0].typename !== "PathItem") {
        alert("사각형 패스를 선택해주세요.");
        return;
    }

    var frontFace = sel[0];
    var bounds = frontFace.geometricBounds; // [left, top, right, bottom]
    var width = bounds[2] - bounds[0];
    var height = bounds[1] - bounds[3];
    var mmToPt = 2.83464567;
    var defaultDepthMm = Math.round(Math.min(width, height) / 2 / mmToPt * 100) / 100;
    var previewItems = [];

    var depthMm = showDepthDialog(defaultDepthMm, function(valueMm) {
        clearPreview();
        previewItems = createCabinet(valueMm * mmToPt);
        app.redraw();
    }, clearPreview);

    clearPreview();
    if (depthMm === null) {
        return;
    }

    createCabinet(depthMm * mmToPt);
    doc.selection = null;

    function createCabinet(depth) {
        frontFace.strokeJoin = StrokeJoin.ROUNDENDJOIN;

        var rightFace = doc.pathItems.add();
        rightFace.setEntirePath([
            [bounds[2], bounds[1]],
            [bounds[2] + depth, bounds[1] + depth],
            [bounds[2] + depth, bounds[3] + depth],
            [bounds[2], bounds[3]]
        ]);
        rightFace.closed = true;
        copyStyle(frontFace, rightFace);

        var topFace = doc.pathItems.add();
        topFace.setEntirePath([
            [bounds[0], bounds[1]],
            [bounds[2], bounds[1]],
            [bounds[2] + depth, bounds[1] + depth],
            [bounds[0] + depth, bounds[1] + depth]
        ]);
        topFace.closed = true;
        copyStyle(frontFace, topFace);

        rightFace.move(frontFace, ElementPlacement.PLACEBEFORE);
        topFace.move(frontFace, ElementPlacement.PLACEBEFORE);

        return [rightFace, topFace];
    }

    function showDepthDialog(defaultValue, onPreview, onClearPreview) {
        var dialog = new Window("dialog", "캐비넷 깊이");
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

    function copyStyle(source, target) {
        target.filled = source.filled;
        if (source.filled) {
            target.fillColor = source.fillColor;
        }

        target.stroked = source.stroked;
        if (source.stroked) {
            target.strokeColor = source.strokeColor;
            target.strokeWidth = source.strokeWidth;
            target.strokeDashes = source.strokeDashes;
            target.strokeCap = source.strokeCap;
            target.strokeJoin = StrokeJoin.ROUNDENDJOIN;
            target.opacity = source.opacity;
        }
    }
})();
