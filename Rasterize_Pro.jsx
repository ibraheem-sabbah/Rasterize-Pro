#target illustrator

(function () {
    if (app.documents.length === 0) {
        alert("Please open an Illustrator document first.");
        return;
    }

    var doc = app.activeDocument;
    var currentSelection = doc.selection;

    if (!currentSelection || currentSelection.length === 0) {
        alert("Please select one or more objects first.");
        return;
    }

    // Store the original selection before processing.
    var selectedRoots = [];
    var i;
    for (i = 0; i < currentSelection.length; i++) {
        selectedRoots.push(currentSelection[i]);
    }

    // ------------------------------------------------------------
    // Persistent User Settings
    // ------------------------------------------------------------
    var resolutions = [2400, 1200, 600, 300, 150, 72];

    var settingsFile = new File(
        Folder.userData.fsName + "/Rasterize_Pro_Settings.txt"
    );

    var savedSettings = {
        resolution: 300,
        background: "transparent",
        ignoreHidden: true
    };

    // Load saved settings.
    try {
        if (settingsFile.exists && settingsFile.open("r")) {
            var settingsContent = settingsFile.read();
            settingsFile.close();

            var lines = settingsContent.split(/\r?\n/);

            for (var p = 0; p < lines.length; p++) {
                var parts = lines[p].split("=");

                if (parts.length < 2) continue;

                var key = parts[0];
                var value = parts.slice(1).join("=");

                if (key === "resolution") {
                    var savedResolution = parseInt(value, 10);

                    for (var ri = 0; ri < resolutions.length; ri++) {
                        if (resolutions[ri] === savedResolution) {
                            savedSettings.resolution = savedResolution;
                            break;
                        }
                    }
                }

                if (key === "background") {
                    if (value === "transparent" || value === "white") {
                        savedSettings.background = value;
                    }
                }

                if (key === "ignoreHidden") {
                    savedSettings.ignoreHidden = (value === "true");
                }
            }
        }
    } catch (settingsError) {
        try {
            if (settingsFile.opened) settingsFile.close();
        } catch (closeSettingsError) {}
    }

    // Save settings.
    function saveRasterizeSettings(resolution, background, ignoreHidden) {
        try {
            if (settingsFile.open("w")) {
                settingsFile.writeln("resolution=" + resolution);
                settingsFile.writeln("background=" + background);
                settingsFile.writeln("ignoreHidden=" + ignoreHidden);
                settingsFile.close();
            }
        } catch (saveError) {
            try {
                if (settingsFile.opened) settingsFile.close();
            } catch (closeSaveError) {}
        }
    }

    // ------------------------------------------------------------
    // Settings Dialog
    // ------------------------------------------------------------
    var dialog = new Window("dialog", "Rasterize Pro");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";

    var resolutionPanel = dialog.add(
        "panel", undefined, "Resolution"
    );

    resolutionPanel.orientation = "column";
    resolutionPanel.alignChildren = "left";
    resolutionPanel.margins = 12;

    var resolutionLabels = [
        "2400 DPI",
        "1200 DPI",
        "600 DPI",
        "300 DPI",
        "150 DPI",
        "72 DPI"
    ];

    var resolutionList = resolutionPanel.add(
        "dropdownlist",
        undefined,
        resolutionLabels
    );

    // Restore previous resolution. First-run default: 300 DPI.
    var defaultIndex = 3;

    for (var rIndex = 0; rIndex < resolutions.length; rIndex++) {
        if (resolutions[rIndex] === savedSettings.resolution) {
            defaultIndex = rIndex;
            break;
        }
    }

    resolutionList.selection = defaultIndex;
    resolutionList.minimumSize.width = 180;

    // Background.
    var backgroundPanel = dialog.add(
        "panel", undefined, "Background"
    );

    backgroundPanel.orientation = "column";
    backgroundPanel.alignChildren = "left";
    backgroundPanel.margins = 12;

    var transparentOption = backgroundPanel.add(
        "radiobutton",
        undefined,
        "Transparent Background"
    );

    var whiteOption = backgroundPanel.add(
        "radiobutton",
        undefined,
        "White Background"
    );

    transparentOption.value =
        savedSettings.background === "transparent";

    whiteOption.value =
        savedSettings.background === "white";

    // Advanced options.
    var hiddenPanel = dialog.add(
        "panel", undefined, "Advanced Options"
    );

    hiddenPanel.orientation = "column";
    hiddenPanel.alignChildren = "left";
    hiddenPanel.margins = 12;

    var ignoreHiddenOption = hiddenPanel.add(
        "checkbox",
        undefined,
        "Ignore Hidden Layers in Bounds"
    );

    ignoreHiddenOption.value = savedSettings.ignoreHidden;
    ignoreHiddenOption.helpTip =
        "ON: Exclude hidden artwork from the calculated image bounds. " +
        "OFF: Include hidden artwork bounds, preserving the corresponding empty space.";

    // Buttons.
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "right";

    buttonGroup.add(
        "button", undefined, "Cancel", { name: "cancel" }
    );

    buttonGroup.add(
        "button", undefined, "Rasterize", { name: "ok" }
    );

    if (dialog.show() !== 1) {
        return;
    }

    // Read selected options.
    var resolution = resolutions[resolutionList.selection.index];
    var useTransparency = transparentOption.value;
    var ignoreHiddenLayers = ignoreHiddenOption.value;

    var selectedBackground =
        useTransparency ? "transparent" : "white";

    // Save settings automatically for the next run.
    saveRasterizeSettings(
        resolution,
        selectedBackground,
        ignoreHiddenLayers
    );

    // ------------------------------------------------------------
    // General Helpers
    // ------------------------------------------------------------
    function getType(item) {
        try {
            return item.typename;
        } catch (e) {
            return "Unknown";
        }
    }

    function getName(item) {
        try {
            if (item.name && item.name !== "") {
                return item.name;
            }
        } catch (e) {}
        return "(unnamed " + getType(item) + ")";
    }

    function sameObject(a, b) {
        try {
            return a === b || a == b;
        } catch (e) {
            return false;
        }
    }

    function hasReference(list, item) {
        for (var n = 0; n < list.length; n++) {
            if (sameObject(list[n], item)) {
                return true;
            }
        }
        return false;
    }

    function getDirectPageItems(container) {
        var result = [];
        var items;

        try {
            items = container.pageItems;
        } catch (e) {
            return result;
        }

        for (var n = 0; n < items.length; n++) {
            var child = items[n];
            try {
                if (sameObject(child.parent, container)) {
                    result.push(child);
                }
            } catch (e2) {}
        }

        return result;
    }

    function isEffectivelyHidden(item) {
        var current = item;
        var depth = 0;

        while (current && depth < 100) {
            var type = getType(current);

            try {
                if (type === "Layer" && current.visible === false) {
                    return true;
                }

                if (type !== "Document" && current.hidden === true) {
                    return true;
                }
            } catch (e) {}

            if (type === "Document") {
                break;
            }

            try {
                current = current.parent;
            } catch (e2) {
                break;
            }

            depth++;
        }

        return false;
    }

    function isGuide(item) {
        var type = getType(item);

        try {
            if (type === "PathItem" && item.guides === true) {
                return true;
            }

            if (type === "CompoundPathItem") {
                for (var p = 0; p < item.pathItems.length; p++) {
                    if (item.pathItems[p].guides === true) {
                        return true;
                    }
                }
            }
        } catch (e) {}

        return false;
    }

    function isClippingObject(item) {
        try {
            if (item.clipping === true) {
                return true;
            }
        } catch (e) {}

        if (getType(item) === "CompoundPathItem") {
            try {
                for (var cp = 0; cp < item.pathItems.length; cp++) {
                    if (item.pathItems[cp].clipping === true) {
                        return true;
                    }
                }
            } catch (e2) {}
        }

        return false;
    }

    // Bounds format: [left, top, right, bottom].
    function readItemBounds(item, useGeometricBounds) {
        var b = null;

        try {
            b = useGeometricBounds ? item.geometricBounds : item.visibleBounds;
        } catch (e) {}

        if (!b || b.length < 4) {
            try {
                b = useGeometricBounds ? item.visibleBounds : item.geometricBounds;
            } catch (e2) {
                return null;
            }
        }

        if (!b || b.length < 4) {
            return null;
        }

        var left = Number(b[0]);
        var top = Number(b[1]);
        var right = Number(b[2]);
        var bottom = Number(b[3]);

        if (!isFinite(left) || !isFinite(top) ||
            !isFinite(right) || !isFinite(bottom)) {
            return null;
        }

        return [
            Math.min(left, right),
            Math.max(top, bottom),
            Math.max(left, right),
            Math.min(top, bottom)
        ];
    }

    function combineBounds(a, b) {
        if (!a) return b;
        if (!b) return a;

        return [
            Math.min(a[0], b[0]),
            Math.max(a[1], b[1]),
            Math.max(a[2], b[2]),
            Math.min(a[3], b[3])
        ];
    }

    function intersectBounds(a, b) {
        if (!a || !b) return null;

        var left = Math.max(a[0], b[0]);
        var top = Math.min(a[1], b[1]);
        var right = Math.min(a[2], b[2]);
        var bottom = Math.max(a[3], b[3]);

        if (right <= left || top <= bottom) {
            return null;
        }

        return [left, top, right, bottom];
    }

    // ------------------------------------------------------------
    // Expand selected Layers into their direct artwork items.
    // Groups remain intact so group effects and clipping appearance
    // are preserved during rasterization.
    // ------------------------------------------------------------
    var targets = [];
    var skippedHiddenLayers = 0;

    function addTarget(item) {
        if (!item) return;

        var type = getType(item);

        if (type === "Layer") {
            addLayerTargets(item);
            return;
        }

        if (!hasReference(targets, item)) {
            targets.push(item);
        }
    }

    function addLayerTargets(layer) {
        if (!layer) return;

        if (isEffectivelyHidden(layer)) {
            skippedHiddenLayers++;
            return;
        }

        var directItems = getDirectPageItems(layer);

        for (var a = 0; a < directItems.length; a++) {
            addTarget(directItems[a]);
        }

        try {
            for (var s = 0; s < layer.layers.length; s++) {
                addLayerTargets(layer.layers[s]);
            }
        } catch (e) {}
    }

    function hasSelectedAncestor(item, selection) {
        var current = null;
        var depth = 0;

        try {
            current = item.parent;
        } catch (e) {
            return false;
        }

        while (current && depth < 100) {
            if (hasReference(selection, current)) {
                return true;
            }

            if (getType(current) === "Document") {
                break;
            }

            try {
                current = current.parent;
            } catch (e2) {
                break;
            }

            depth++;
        }

        return false;
    }

    for (i = 0; i < selectedRoots.length; i++) {
        if (hasSelectedAncestor(selectedRoots[i], selectedRoots)) {
            continue;
        }

        addTarget(selectedRoots[i]);
    }

    if (targets.length === 0) {
        alert("No selected artwork was found to rasterize.");
        return;
    }

    // ------------------------------------------------------------
    // Smart Visible Bounds
    // ON: hidden artwork is excluded from calculated bounds.
    // OFF: hidden artwork bounds are included, preserving empty space.
    // ------------------------------------------------------------
    var hiddenObjectsExcluded = 0;
    var guidesExcluded = 0;
    var clippingMaskFallbacks = 0;
    var details = [];

    function addDetail(message) {
        if (details.length < 25) {
            details.push(message);
        }
    }

    // Find a clipping path in a clipped group. Illustrator may expose
    // the clipping object directly, or inside a nested group/compound path.
    function findClippingMask(container, depth) {
        depth = depth || 0;
        if (!container || depth > 40) return null;

        var children = getDirectPageItems(container);
        var c;

        // Prefer a direct clipping object: this is the usual Illustrator structure.
        for (c = 0; c < children.length; c++) {
            if (isClippingObject(children[c])) {
                return children[c];
            }
        }

        // Fallback: search nested groups when the clipping path is not a direct child.
        for (c = 0; c < children.length; c++) {
            var childType = getType(children[c]);
            if (childType === "GroupItem" || childType === "Layer") {
                var nestedMask = findClippingMask(children[c], depth + 1);
                if (nestedMask) return nestedMask;
            }
        }

        return null;
    }

    function getSmartVisibleBounds(item, depth, excludedMask) {
        depth = depth || 0;

        if (!item || depth > 80) {
            return null;
        }

        // Exclude the clipping path from artwork bounds. Its own bounds
        // are applied separately as the clipping boundary.
        if (excludedMask && sameObject(item, excludedMask)) {
            return null;
        }

        // This is the only hidden-object condition.
        // ON: hidden items do not enlarge the result.
        // OFF: hidden items contribute their bounds, preserving empty space.
        if (isEffectivelyHidden(item) && ignoreHiddenLayers) {
            hiddenObjectsExcluded++;
            addDetail(
                "Hidden artwork excluded from bounds: " + getName(item)
            );
            return null;
        }

        if (isGuide(item)) {
            guidesExcluded++;
            addDetail("Guide excluded: " + getName(item));
            return null;
        }

        var type = getType(item);

        // Layers are measured by their direct contents.
        if (type === "Layer") {
            var layerBounds = null;
            var layerItems = getDirectPageItems(item);

            for (var li = 0; li < layerItems.length; li++) {
                layerBounds = combineBounds(
                    layerBounds,
                    getSmartVisibleBounds(layerItems[li], depth + 1, excludedMask)
                );
            }

            try {
                for (var sl = 0; sl < item.layers.length; sl++) {
                    layerBounds = combineBounds(
                        layerBounds,
                        getSmartVisibleBounds(item.layers[sl], depth + 1, excludedMask)
                    );
                }
            } catch (e) {}

            return layerBounds;
        }

        if (type === "GroupItem") {
            var children = getDirectPageItems(item);
            var clipped = false;

            try {
                clipped = item.clipped === true;
            } catch (e2) {}

            if (clipped) {
                var clippingItem = findClippingMask(item, 0);

                if (clippingItem) {
                    // Use the clipping path's geometric bounds as the mask limit.
                    var maskBounds = readItemBounds(clippingItem, true);
                    var clippedContentBounds = null;

                    // Recurse through every child, including nested groups,
                    // while excluding the clipping object itself.
                    for (var k = 0; k < children.length; k++) {
                        clippedContentBounds = combineBounds(
                            clippedContentBounds,
                            getSmartVisibleBounds(
                                children[k],
                                depth + 1,
                                clippingItem
                            )
                        );
                    }

                    if (maskBounds && clippedContentBounds) {
                        return intersectBounds(clippedContentBounds, maskBounds);
                    }

                    if (maskBounds && !clippedContentBounds) {
                        return null;
                    }
                } else {
                    // Report an unusual clipping-group structure. Continue with
                    // descendant bounds, allowing hidden-item filtering to work.
                    clippingMaskFallbacks++;
                    addDetail(
                        "Clipping path not found directly or in nested groups: " +
                        getName(item)
                    );
                }
            }

            // Regular group, or fallback when Illustrator does not expose
            // the clipping path through the scripting DOM.
            var groupBounds = null;

            for (var g = 0; g < children.length; g++) {
                groupBounds = combineBounds(
                    groupBounds,
                    getSmartVisibleBounds(children[g], depth + 1, excludedMask)
                );
            }

            return groupBounds;
        }

        // Ordinary path, text, raster, placed, symbol and other items.
        return readItemBounds(item, false);
    }

    // ------------------------------------------------------------
    // Temporary unlock helpers. Original lock states are restored.
    // ------------------------------------------------------------
    function unlockItemAndParents(item) {
        var records = [];
        var current = item;
        var depth = 0;

        while (current && depth < 100) {
            try {
                if (current.locked === true) {
                    records.push({
                        object: current,
                        wasLocked: true
                    });
                }
            } catch (e) {}

            if (getType(current) === "Document") {
                break;
            }

            try {
                current = current.parent;
            } catch (e2) {
                break;
            }

            depth++;
        }

        // Unlock from outermost parent to selected item.
        for (var r = records.length - 1; r >= 0; r--) {
            try {
                records[r].object.locked = false;
            } catch (e3) {}
        }

        return records;
    }

    function restoreLocks(records, originalItem, replacementItem) {
        for (var r = 0; r < records.length; r++) {
            try {
                if (sameObject(records[r].object, originalItem) && replacementItem) {
                    replacementItem.locked = records[r].wasLocked;
                } else {
                    records[r].object.locked = records[r].wasLocked;
                }
            } catch (e) {}
        }
    }

    // ------------------------------------------------------------
    // Rasterization
    // ------------------------------------------------------------
    var converted = 0;
    var skipped = 0;
    var failed = 0;
    var outputItems = [];

    var options = new RasterizeOptions();
    options.resolution = resolution;
    options.transparency = useTransparency;
    options.backgroundBlack = false;
    options.clippingMask = false;
    options.convertSpotColors = false;

    for (var j = targets.length - 1; j >= 0; j--) {
        var item = targets[j];
        var itemName = getName(item);
        var itemType = getType(item);

        var smartBounds = getSmartVisibleBounds(item, 0);

        if (!smartBounds) {
            skipped++;

            addDetail(
                "Skipped (no artwork bounds): " +
                itemName + " [" + itemType + "]"
            );

            continue;
        }

        var originalLeft = smartBounds[0];
        var originalTop = smartBounds[1];
        var originalWidth = smartBounds[2] - smartBounds[0];
        var originalHeight = smartBounds[1] - smartBounds[3];

        if (originalWidth <= 0 || originalHeight <= 0) {
            skipped++;

            addDetail(
                "Skipped (zero-size bounds): " +
                itemName + " [" + itemType + "]"
            );

            continue;
        }

        var rasterBounds = [
            smartBounds[0],
            smartBounds[1],
            smartBounds[2],
            smartBounds[3]
        ];

        var lockRecords = unlockItemAndParents(item);

        try {
            if (itemType === "TextFrame") {
                options.antiAliasingMethod =
                    AntiAliasingMethod.TYPEOPTIMIZED;
            } else {
                options.antiAliasingMethod =
                    AntiAliasingMethod.ARTOPTIMIZED;
            }

            var rasterResult = doc.rasterize(
                item,
                rasterBounds,
                options
            );

            if (!rasterResult) {
                throw new Error(
                    "Illustrator did not return a raster image."
                );
            }

            // Restore the calculated bounds and position.
            rasterResult.width = originalWidth;
            rasterResult.height = originalHeight;
            rasterResult.position = [
                originalLeft,
                originalTop
            ];

            restoreLocks(
                lockRecords,
                item,
                rasterResult
            );

            outputItems.push(rasterResult);
            converted++;

        } catch (e4) {
            restoreLocks(
                lockRecords,
                item,
                null
            );

            failed++;

            var errorText = "Unknown error";

            try {
                errorText = e4.message || e4.toString();
            } catch (e5) {}

            addDetail(
                "Failed: " + itemName + " [" +
                itemType + "] - " + errorText
            );
        }
    }

    // Select resulting raster images for easy review.
    if (outputItems.length > 0) {
        try {
            doc.selection = outputItems;
        } catch (selectionError) {}
    }

    // ------------------------------------------------------------
    // Result Summary
    // ------------------------------------------------------------
    var result = "Rasterization Complete\n\n";

    result += "Resolution: " + resolution + " DPI\n";
    result += "Background: " +
        (useTransparency ? "Transparent" : "White") + "\n";

    result += "Ignore Hidden Layers in Bounds: " +
        (ignoreHiddenLayers ? "ON" : "OFF") + "\n";

    result += "Selected items: " + targets.length + "\n";
    result += "Successfully rasterized: " + converted + "\n";
    result += "Skipped: " + skipped + "\n";
    result += "Failed: " + failed + "\n";

    result += "Hidden objects excluded from bounds: " +
        hiddenObjectsExcluded + "\n";

    result += "Guides excluded from bounds: " +
        guidesExcluded + "\n";
    result += "Clipping groups using fallback bounds: " +
        clippingMaskFallbacks + "\n";

    if (skippedHiddenLayers > 0) {
        result += "Hidden layers skipped: " +
            skippedHiddenLayers + "\n";
    }

    if (details.length > 0) {
        result += "\nDetails:\n";

        for (var d = 0; d < details.length; d++) {
            result += "- " + details[d] + "\n";
        }

        if (details.length >= 25) {
            result += "Additional details were omitted.\n";
        }
    }

    result +=
        "\nThis Script Made By Ibraheem Sabbah - www.linkedin.com/in/ibraheem-sabbah";

    alert(result);
})();
