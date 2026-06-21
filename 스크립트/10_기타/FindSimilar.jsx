/*
  FindSimilar.jsx
  기능: 선택한 기준 개체와 비슷한 개체를 조건별로 찾아 선택하고,
        찾은 개체들을 교체본 개체로 일괄 교체합니다.
  특징: 비modal 패널(palette)로 떠서 찾기/교체를 반복 실행할 수 있습니다.
*/

#target illustrator
#targetengine "findsimilar"

(function () {
    // 재실행 시 기존 패널을 다시 보여줍니다(중복 창 방지).
    if ($.global.__findSimilarWin) {
        try {
            $.global.__findSimilarWin.show();
            return;
        } catch (eShow) {
            $.global.__findSimilarWin = null;
        }
    }

    var win = buildPalette();
    $.global.__findSimilarWin = win;
    win.show();

    // ===================== UI =====================

    function buildPalette() {
        var settings = loadSettings();
        var foundItems = []; // 직전 찾기 결과(교체 대상 후보). 패널이 살아있는 동안 유지됩니다.

        var w = new Window("palette", "Find Similar");
        w.orientation = "column";
        w.alignChildren = "fill";
        w.margins = 16;

        var scopePanel = w.add("panel", undefined, "검색 범위");
        scopePanel.orientation = "row";
        scopePanel.alignChildren = "left";
        scopePanel.margins = 12;
        var rbDoc = scopePanel.add("radiobutton", undefined, "문서 전체");
        var rbSelection = scopePanel.add("radiobutton", undefined, "현재 선택 안");
        rbDoc.value = settings.scope !== "selection";
        rbSelection.value = settings.scope === "selection";

        var conditionPanel = w.add("panel", undefined, "조건");
        conditionPanel.orientation = "column";
        conditionPanel.alignChildren = "left";
        conditionPanel.margins = 12;

        var row1 = conditionPanel.add("group");
        row1.orientation = "row";
        var chkGeometry = row1.add("checkbox", undefined, "Geometry");
        var chkObjectType = row1.add("checkbox", undefined, "Object Type");
        var chkSize = row1.add("checkbox", undefined, "Size");

        var row2 = conditionPanel.add("group");
        row2.orientation = "row";
        var chkFill = row2.add("checkbox", undefined, "Fill");
        var chkStroke = row2.add("checkbox", undefined, "Stroke");
        var chkStrokeWidth = row2.add("checkbox", undefined, "Stroke Width");
        var chkOpacity = row2.add("checkbox", undefined, "Opacity");

        chkGeometry.value = settings.geometry;
        chkObjectType.value = settings.objectType;
        chkFill.value = settings.fill;
        chkStroke.value = settings.stroke;
        chkStrokeWidth.value = settings.strokeWidth;
        chkSize.value = settings.size;
        chkOpacity.value = settings.opacity;

        var geoPanel = w.add("panel", undefined, "Geometry 옵션");
        geoPanel.orientation = "column";
        geoPanel.alignChildren = "left";
        geoPanel.margins = 12;

        var geoRow1 = geoPanel.add("group");
        geoRow1.orientation = "row";
        var chkScaleAllowed = geoRow1.add("checkbox", undefined, "Scale allowed");
        var chkRotationAllowed = geoRow1.add("checkbox", undefined, "Rotation allowed");
        var chkMirrorAllowed = geoRow1.add("checkbox", undefined, "Mirror allowed");

        chkScaleAllowed.value = settings.scaleAllowed;
        chkRotationAllowed.value = settings.rotationAllowed;
        chkMirrorAllowed.value = settings.mirrorAllowed;

        var tolPanel = w.add("panel", undefined, "Tolerance");
        tolPanel.orientation = "column";
        tolPanel.alignChildren = "left";
        tolPanel.margins = 12;

        var posRow = tolPanel.add("group");
        posRow.orientation = "row";
        posRow.add("statictext", undefined, "위치:");
        var inputPosTol = posRow.add("edittext", undefined, String(settings.tolerance));
        inputPosTol.characters = 6;
        posRow.add("statictext", undefined, "pt (size / geometry / stroke width / opacity)");

        var colRow = tolPanel.add("group");
        colRow.orientation = "row";
        colRow.add("statictext", undefined, "색상:");
        var inputColTol = colRow.add("edittext", undefined, String(settings.colorTolerance));
        inputColTol.characters = 6;
        colRow.add("statictext", undefined, "% (fill / stroke)");

        var replacePanel = w.add("panel", undefined, "교체");
        replacePanel.orientation = "row";
        replacePanel.alignChildren = "left";
        replacePanel.margins = 12;
        var chkFit = replacePanel.add("checkbox", undefined, "크기 맞춤(fit)");
        chkFit.value = settings.fit;

        var statusText = w.add("statictext", undefined, "기준 개체 선택 후 [찾기]를 누르세요.");
        statusText.characters = 42;

        var btns = w.add("group");
        btns.alignment = "center";
        var btnFind = btns.add("button", undefined, "찾기");
        var btnReplace = btns.add("button", undefined, "교체");
        var btnClose = btns.add("button", undefined, "닫기");

        btnFind.onClick = function () { doFind(); };
        btnReplace.onClick = function () { doReplace(); };
        btnClose.onClick = function () { w.hide(); };

        w.onClose = function () {
            try { saveSettings(readOptions()); } catch (e) {}
            return true;
        };

        function readOptions() {
            var posTol = parseFloat(inputPosTol.text);
            if (isNaN(posTol) || posTol < 0) posTol = 0.5;
            var colTol = parseFloat(inputColTol.text);
            if (isNaN(colTol) || colTol < 0) colTol = 5;
            return {
                scope: rbSelection.value ? "selection" : "document",
                geometry: chkGeometry.value,
                objectType: chkObjectType.value,
                fill: chkFill.value,
                stroke: chkStroke.value,
                strokeWidth: chkStrokeWidth.value,
                size: chkSize.value,
                opacity: chkOpacity.value,
                scaleAllowed: chkScaleAllowed.value,
                rotationAllowed: chkRotationAllowed.value,
                mirrorAllowed: chkMirrorAllowed.value,
                tolerance: posTol,
                colorTolerance: colTol,
                fit: chkFit.value
            };
        }

        function doFind() {
            if (app.documents.length === 0) { alert("열린 문서가 없습니다."); return; }
            var doc = app.activeDocument;
            var sel = doc.selection;
            if (!sel || sel.length < 1) {
                alert("기준 개체를 1개 선택한 뒤 [찾기]를 누르세요.");
                return;
            }

            var reference = sel[0];
            var options = readOptions();
            saveSettings(options);

            var candidates = options.scope === "selection"
                ? selectionCandidates(sel, reference)
                : documentCandidates(doc, reference);

            var refProfile = buildProfile(reference, options);
            var matches = [reference];

            for (var i = 0; i < candidates.length; i++) {
                try {
                    if (candidateMatches(refProfile, candidates[i], options)) {
                        matches.push(candidates[i]);
                    }
                } catch (e) {}
            }

            doc.selection = null;
            for (var m = 0; m < matches.length; m++) {
                try { matches[m].selected = true; } catch (e2) {}
            }
            foundItems = matches;
            app.redraw();
            statusText.text = matches.length + "개 선택됨(기준 포함). 교체본 선택 후 [교체].";
        }

        function doReplace() {
            if (app.documents.length === 0) { alert("열린 문서가 없습니다."); return; }
            var doc = app.activeDocument;

            if (!foundItems || foundItems.length === 0) {
                alert("먼저 [찾기]를 실행하세요.");
                return;
            }

            var sel = doc.selection;
            if (!sel || sel.length < 1) {
                alert("교체본 개체를 1개 선택한 뒤 [교체]를 누르세요.");
                return;
            }
            var replacement = sel[0];
            var options = readOptions();

            // 교체본이 찾은 개체에 포함되면 모호하므로 중단합니다.
            var targets = [];
            for (var i = 0; i < foundItems.length; i++) {
                if (foundItems[i] === replacement) {
                    alert("교체본이 찾은 개체에 포함됩니다. 교체본은 따로 선택하세요.");
                    return;
                }
                if (isAlive(foundItems[i])) targets.push(foundItems[i]);
            }

            if (targets.length === 0) {
                alert("교체할 대상이 없습니다.");
                return;
            }

            var newItems = [];
            for (var t = 0; t < targets.length; t++) {
                try {
                    var target = targets[t];
                    var dup = replacement.duplicate(target, ElementPlacement.PLACEBEFORE);
                    if (options.fit) fitToTarget(dup, target);
                    centerOn(dup, target);
                    target.remove();
                    newItems.push(dup);
                } catch (e3) {}
            }

            foundItems = [];
            doc.selection = null;
            for (var n = 0; n < newItems.length; n++) {
                try { newItems[n].selected = true; } catch (e4) {}
            }
            app.redraw();
            statusText.text = newItems.length + "개 교체됨.";
        }

        return w;
    }

    // ===================== 후보 수집 =====================

    function selectionCandidates(selectionItems, refItem) {
        var result = [];
        for (var i = 0; i < selectionItems.length; i++) {
            if (selectionItems[i] !== refItem && isUsableItem(selectionItems[i])) {
                result.push(selectionItems[i]);
            }
        }
        return result;
    }

    function documentCandidates(document, refItem) {
        var result = [];
        for (var i = 0; i < document.pageItems.length; i++) {
            var item = document.pageItems[i];
            if (item === refItem) continue;
            if (!isTopLevelItem(item)) continue;
            if (!isUsableItem(item)) continue;
            result.push(item);
        }
        return result;
    }

    function isTopLevelItem(item) {
        return item.parent && item.parent.typename !== "GroupItem" && item.parent.typename !== "CompoundPathItem";
    }

    function isUsableItem(item) {
        if (!item || item.guides) return false;
        var current = item;
        while (current && current.typename !== "Document") {
            if (current.locked || current.hidden || current.visible === false) return false;
            current = current.parent;
        }
        return true;
    }

    function isAlive(item) {
        try {
            var t = item.typename;
            return !!t;
        } catch (e) {
            return false;
        }
    }

    // ===================== 매칭 (성능: 기준 프로필 1회 선계산) =====================

    function buildProfile(item, opts) {
        var paths = collectPathItems(item);
        return {
            typename: item.typename,
            opacity: item.opacity,
            bounds: safeBounds(item),
            fill: opts.fill ? getStyleListFrom(paths, "fill") : null,
            stroke: opts.stroke ? getStyleListFrom(paths, "stroke") : null,
            strokeWidth: opts.strokeWidth ? getStrokeWidthListFrom(paths) : null,
            geometry: opts.geometry ? sortedGeometryFrom(paths, opts) : null
        };
    }

    // 싼 검사(type/size/opacity)를 먼저 하고, 통과했을 때만 비싼 collectPathItems를 호출합니다.
    function candidateMatches(refProfile, testItem, opts) {
        if (opts.objectType && refProfile.typename !== testItem.typename) return false;
        if (opts.size && !opts.scaleAllowed && !sizeMatchesBounds(refProfile.bounds, safeBounds(testItem), opts.tolerance)) return false;
        if (opts.opacity && Math.abs(refProfile.opacity - testItem.opacity) > opts.tolerance) return false;

        if (opts.fill || opts.stroke || opts.strokeWidth || opts.geometry) {
            var paths = collectPathItems(testItem);
            if (opts.fill && !styleListMatches(refProfile.fill, getStyleListFrom(paths, "fill"), opts.colorTolerance)) return false;
            if (opts.stroke && !styleListMatches(refProfile.stroke, getStyleListFrom(paths, "stroke"), opts.colorTolerance)) return false;
            if (opts.strokeWidth && !strokeWidthListMatches(refProfile.strokeWidth, getStrokeWidthListFrom(paths), opts.tolerance)) return false;
            if (opts.geometry && !geometryMatchesSorted(refProfile.geometry, sortedGeometryFrom(paths, opts), opts)) return false;
        }
        return true;
    }

    function sizeMatchesBounds(ab, bb, tol) {
        if (!ab || !bb) return false;
        var aw = Math.abs(ab[2] - ab[0]);
        var ah = Math.abs(ab[1] - ab[3]);
        var bw = Math.abs(bb[2] - bb[0]);
        var bh = Math.abs(bb[1] - bb[3]);
        return Math.abs(aw - bw) <= tol && Math.abs(ah - bh) <= tol;
    }

    function safeBounds(item) {
        try {
            return item.geometricBounds;
        } catch (e) {
            try {
                return item.visibleBounds;
            } catch (e2) {
                return null;
            }
        }
    }

    // ===================== 경로 수집 =====================

    function collectPathItems(item) {
        var result = [];
        collectPathItemsInto(item, result);
        return result;
    }

    function collectPathItemsInto(item, result) {
        if (!item) return;
        if (item.typename === "PathItem") {
            if (!item.guides && !item.clipping) result.push(item);
            return;
        }
        if (item.typename === "CompoundPathItem") {
            for (var i = 0; i < item.pathItems.length; i++) {
                if (!item.pathItems[i].guides && !item.pathItems[i].clipping) result.push(item.pathItems[i]);
            }
            return;
        }
        if (item.typename === "GroupItem") {
            for (var g = 0; g < item.pageItems.length; g++) {
                collectPathItemsInto(item.pageItems[g], result);
            }
            return;
        }
        if (item.typename === "TextFrame") {
            var temp = null;
            var outline = null;
            try {
                temp = item.duplicate();
                outline = temp.createOutline();
                collectPathItemsInto(outline, result);
            } catch (e) {
            } finally {
                try {
                    if (outline) outline.remove();
                    if (temp) temp.remove();
                } catch (removeError) {}
            }
        }
    }

    // ===================== Fill / Stroke 색상 =====================

    function getStyleListFrom(paths, mode) {
        var result = [];
        for (var i = 0; i < paths.length; i++) {
            if (mode === "fill") {
                result.push(paths[i].filled ? colorKey(paths[i].fillColor) : "NoFill");
            } else {
                result.push(paths[i].stroked ? colorKey(paths[i].strokeColor) : "NoStroke");
            }
        }
        return result.sort();
    }

    function colorKey(color) {
        if (!color) return "None";
        try {
            if (color.typename === "RGBColor") {
                return "RGB:" + round(color.red, 2) + "," + round(color.green, 2) + "," + round(color.blue, 2);
            }
            if (color.typename === "CMYKColor") {
                return "CMYK:" + round(color.cyan, 2) + "," + round(color.magenta, 2) + "," + round(color.yellow, 2) + "," + round(color.black, 2);
            }
            if (color.typename === "GrayColor") {
                return "Gray:" + round(color.gray, 2);
            }
            if (color.typename === "NoColor") {
                return "NoColor";
            }
            return color.typename;
        } catch (e) {
            return "Unknown";
        }
    }

    function styleListMatches(a, b, colorTol) {
        if (!a || !b || a.length !== b.length) return false;
        for (var i = 0; i < a.length; i++) {
            if (!colorsMatch(a[i], b[i], colorTol)) return false;
        }
        return true;
    }

    // colorTol은 % 단위입니다. 채널차를 채널 범위 대비 %로 환산해 비교합니다.
    function colorsMatch(a, b, colorTol) {
        if (a === b) return true;
        var pa = parseColorKey(a);
        var pb = parseColorKey(b);
        if (!pa || !pb || pa.type !== pb.type || pa.values.length !== pb.values.length) return false;
        var scale = (pa.type === "RGB") ? (100 / 255) : 1; // CMYK / Gray는 이미 0~100
        for (var i = 0; i < pa.values.length; i++) {
            if (Math.abs(pa.values[i] - pb.values[i]) * scale > colorTol) return false;
        }
        return true;
    }

    function parseColorKey(key) {
        var parts = String(key).split(":");
        if (parts.length !== 2) return null;
        var nums = parts[1].split(",");
        var values = [];
        for (var i = 0; i < nums.length; i++) {
            var n = parseFloat(nums[i]);
            if (isNaN(n)) return null;
            values.push(n);
        }
        return { type: parts[0], values: values };
    }

    // ===================== Stroke Width =====================

    function getStrokeWidthListFrom(paths) {
        var result = [];
        for (var i = 0; i < paths.length; i++) {
            result.push(paths[i].stroked ? paths[i].strokeWidth : 0);
        }
        return result.sort(numericSort);
    }

    function strokeWidthListMatches(a, b, tol) {
        if (!a || !b || a.length !== b.length) return false;
        for (var i = 0; i < a.length; i++) {
            if (Math.abs(a[i] - b[i]) > tol) return false;
        }
        return true;
    }

    // ===================== Geometry (베지어 핸들 포함) =====================

    function sortedGeometryFrom(paths, opts) {
        var result = [];
        for (var i = 0; i < paths.length; i++) {
            result.push(pathSignature(paths[i], opts));
        }
        result.sort(signatureSort);
        return result;
    }

    function pathSignature(pathItem, opts) {
        var pts = pathItem.pathPoints;
        var raw = [];
        for (var i = 0; i < pts.length; i++) {
            var p = pts[i];
            raw.push({
                anchor: [p.anchor[0], p.anchor[1]],
                left: [p.leftDirection[0], p.leftDirection[1]],
                right: [p.rightDirection[0], p.rightDirection[1]]
            });
        }
        return {
            closed: pathItem.closed,
            count: raw.length,
            points: normalizePointSet(raw, opts)
        };
    }

    // 중심이동/회전/스케일을 anchor 기준으로 계산하고, 핸들에도 동일 변환을 적용합니다.
    function normalizePointSet(pts, opts) {
        if (!pts.length) return [];

        var cx = 0, cy = 0;
        for (var i = 0; i < pts.length; i++) {
            cx += pts[i].anchor[0];
            cy += pts[i].anchor[1];
        }
        cx /= pts.length;
        cy /= pts.length;

        var shifted = [];
        var maxX = 0, maxY = 0;
        for (var j = 0; j < pts.length; j++) {
            var a = [pts[j].anchor[0] - cx, pts[j].anchor[1] - cy];
            var l = [pts[j].left[0] - cx, pts[j].left[1] - cy];
            var r = [pts[j].right[0] - cx, pts[j].right[1] - cy];
            shifted.push({ anchor: a, left: l, right: r });
            if (Math.abs(a[0]) > maxX) maxX = Math.abs(a[0]);
            if (Math.abs(a[1]) > maxY) maxY = Math.abs(a[1]);
        }

        if (opts.rotationAllowed) {
            var first = firstNonZeroAnchor(shifted);
            var angle = Math.atan2(first[1], first[0]);
            rotatePointSet(shifted, -angle);
        }

        var scaleX = opts.scaleAllowed ? (maxX || 1) : 1;
        var scaleY = opts.scaleAllowed ? (maxY || 1) : 1;
        var normalized = [];
        for (var k = 0; k < shifted.length; k++) {
            normalized.push({
                anchor: [shifted[k].anchor[0] / scaleX, shifted[k].anchor[1] / scaleY],
                left: [shifted[k].left[0] / scaleX, shifted[k].left[1] / scaleY],
                right: [shifted[k].right[0] / scaleX, shifted[k].right[1] / scaleY]
            });
        }
        return normalized;
    }

    function firstNonZeroAnchor(pts) {
        for (var i = 0; i < pts.length; i++) {
            if (Math.abs(pts[i].anchor[0]) > 0.0001 || Math.abs(pts[i].anchor[1]) > 0.0001) return pts[i].anchor;
        }
        return [1, 0];
    }

    function rotatePointSet(pts, angle) {
        var cos = Math.cos(angle);
        var sin = Math.sin(angle);
        for (var i = 0; i < pts.length; i++) {
            pts[i].anchor = rotateCoord(pts[i].anchor, cos, sin);
            pts[i].left = rotateCoord(pts[i].left, cos, sin);
            pts[i].right = rotateCoord(pts[i].right, cos, sin);
        }
    }

    function rotateCoord(p, cos, sin) {
        return [p[0] * cos - p[1] * sin, p[0] * sin + p[1] * cos];
    }

    function geometryMatchesSorted(refGeo, testGeo, opts) {
        if (!refGeo || !testGeo || refGeo.length !== testGeo.length) return false;
        for (var i = 0; i < refGeo.length; i++) {
            if (!pathSignatureMatches(refGeo[i], testGeo[i], opts)) return false;
        }
        return true;
    }

    function pathSignatureMatches(a, b, opts) {
        if (a.closed !== b.closed || a.count !== b.count) return false;
        if (pointSetMatches(a.points, b.points, opts.tolerance, false)) return true;
        if (opts.mirrorAllowed && pointSetMatches(a.points, b.points, opts.tolerance, true)) return true;
        return false;
    }

    function pointSetMatches(a, b, tol, mirror) {
        if (a.length !== b.length) return false;
        for (var i = 0; i < a.length; i++) {
            var ba = mirror ? [-b[i].anchor[0], b[i].anchor[1]] : b[i].anchor;
            var bl = mirror ? [-b[i].left[0], b[i].left[1]] : b[i].left;
            var br = mirror ? [-b[i].right[0], b[i].right[1]] : b[i].right;
            if (!coordClose(a[i].anchor, ba, tol)) return false;
            if (!coordClose(a[i].left, bl, tol)) return false;
            if (!coordClose(a[i].right, br, tol)) return false;
        }
        return true;
    }

    function coordClose(p, q, tol) {
        return Math.abs(p[0] - q[0]) <= tol && Math.abs(p[1] - q[1]) <= tol;
    }

    // ===================== 교체 보조 =====================

    function fitToTarget(dup, target) {
        var dw = dup.width;
        var dh = dup.height;
        if (dw <= 0 || dh <= 0) return;
        var ratio = Math.min(target.width / dw, target.height / dh);
        if (!(ratio > 0) || !isFinite(ratio)) return;
        var pct = ratio * 100;
        dup.resize(pct, pct, true, true, true, true, pct, Transformation.CENTER);
    }

    function centerOn(dup, target) {
        dup.left = target.left + (target.width - dup.width) / 2;
        dup.top = target.top - (target.height - dup.height) / 2;
    }

    // ===================== 공통 유틸 =====================

    function round(value, digits) {
        var m = Math.pow(10, digits);
        return Math.round(value * m) / m;
    }

    function numericSort(a, b) {
        return a - b;
    }

    function signatureSort(a, b) {
        if (a.count !== b.count) return a.count - b.count;
        if (a.closed !== b.closed) return a.closed ? -1 : 1;
        return 0;
    }

    // ===================== 설정 저장/불러오기 =====================

    function defaultSettings() {
        return {
            scope: "document",
            geometry: true,
            objectType: true,
            fill: true,
            stroke: true,
            strokeWidth: true,
            size: true,
            opacity: false,
            scaleAllowed: false,
            rotationAllowed: false,
            mirrorAllowed: false,
            tolerance: 0.5,
            colorTolerance: 5,
            fit: false
        };
    }

    function settingsFile() {
        var folder = new Folder(Folder.myDocuments + "/Adobe Scripts");
        if (!folder.exists) folder.create();
        return new File(folder.fsName + "/FindSimilar_settings.json");
    }

    function loadSettings() {
        var defaults = defaultSettings();
        var file = settingsFile();
        if (!file.exists) return defaults;

        try {
            file.open("r");
            var text = file.read();
            file.close();
            var loaded = JSON.parse(text);
            for (var key in defaults) {
                if (defaults.hasOwnProperty(key) && loaded.hasOwnProperty(key)) {
                    defaults[key] = loaded[key];
                }
            }
        } catch (e) {
            try { file.close(); } catch (closeError) {}
        }
        return defaults;
    }

    function saveSettings(options) {
        var file = settingsFile();
        try {
            file.open("w");
            file.write(JSON.stringify({
                scope: options.scope,
                geometry: options.geometry,
                objectType: options.objectType,
                fill: options.fill,
                stroke: options.stroke,
                strokeWidth: options.strokeWidth,
                size: options.size,
                opacity: options.opacity,
                scaleAllowed: options.scaleAllowed,
                rotationAllowed: options.rotationAllowed,
                mirrorAllowed: options.mirrorAllowed,
                tolerance: options.tolerance,
                colorTolerance: options.colorTolerance,
                fit: options.fit
            }, null, 2));
            file.close();
        } catch (e) {
            try { file.close(); } catch (closeError) {}
        }
    }
})();
