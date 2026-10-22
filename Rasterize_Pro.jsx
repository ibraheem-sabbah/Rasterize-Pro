
#target illustrator

(function () {
    if (app.documents.length === 0) {
        alert("Please open an Illustrator document first.");
        return;
    }

    var doc = app.activeDocument;
    var originalSelection = doc.selection;

    if (!originalSelection || originalSelection.length === 0) {
        alert("Please select one or more objects first.");
        return;
    }

    var selectedItems = [];
    var i;

    for (i = 0; i < originalSelection.length; i++) {
        selectedItems.push(originalSelection[i]);
    }

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

    function hasReference(list, item) {
        var n;
        for (n = 0; n < list.length; n++) {
            if (list[n] === item) {
                return true;
            }
        }
        return false;
    }

    function hasSelectedAncestor(item, selection) {
        var parentItem = null;
        var depth = 0;

        try {
            parentItem = item.parent;
        } catch (e) {
            return false;
        }

        while (parentItem && depth < 100) {
            if (hasReference(selection, parentItem)) {
                return true;
            }

            if (getType(parentItem) === "Document") {
                break;
            }

            try {
                parentItem = parentItem.parent;
            } catch (e2) {
                break;
            }

            depth++;
        }

        return false;
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

    var dialog = new Window("dialog", "Rasterize Selected Items");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";

    var resolutionPanel = dialog.add(
        "panel", undefined, "Resolution"
    );
    resolutionPanel.orientation = "column";
    resolutionPanel.alignChildren = "left";
    resolutionPanel.margins = 12;

    
        var resolutionList = resolutionPanel.add(
    "dropdownlist",
    undefined,
    [
        "2400 DPI",
        "1200 DPI",
        "600 DPI",
        "300 DPI",
        "150 DPI",
        "72 DPI"
    ]
);

    resolutionList.selection = 0;
    resolutionList.minimumSize.width = 180;

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

    transparentOption.value = true;
    whiteOption.value = false;

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

    var resolutions = [2400, 1200, 600, 300, 150, 72];
    var resolution = resolutions[resolutionList.selection.index];
    var useTransparency = transparentOption.value;

    var targets = [];
    var skipped = 0;
    var converted = 0;
    var failed = 0;
    var details = [];
    var outputItems = [];

    for (i = 0; i < selectedItems.length; i++) {
        var candidate = selectedItems[i];

        if (hasSelectedAncestor(candidate, selectedItems)) {
            continue;
        }

        if (!hasReference(targets, candidate)) {
            targets.push(candidate);
        }
    }

    function addDetail(message) {
        if (details.length < 30) {
            details.push(message);
        }
    }

    for (var j = 0; j < targets.length; j++) {
        var item = targets[j];
        var itemName = getName(item);
        var itemType = getType(item);

        if (isEffectivelyHidden(item)) {
            skipped++;
            addDetail(
                "Skipped (hidden): " + itemName + " [" + itemType + "]"
            );
            continue;
        }

        if (isGuide(item)) {
            skipped++;
            addDetail(
                "Skipped (guide): " + itemName + " [" + itemType + "]"
            );
            continue;
        }

        var wasLocked = false;
        var lockStateKnown = false;

        try {
            wasLocked = item.locked === true;
            lockStateKnown = true;
        } catch (e) {}

        try {
            if (wasLocked) {
                item.locked = false;
            }

            var bounds = item.visibleBounds;

            if (!bounds || bounds.length < 4) {
                throw new Error("Could not read object bounds.");
            }

           var originalLeft = bounds[0];
           var originalTop = bounds[1];
           var originalWidth = bounds[2] - bounds[0];
           var originalHeight = bounds[1] - bounds[3];

           var rasterBounds = [
               originalLeft,
               originalTop,
               bounds[2],
               bounds[3]
            ];

            var options = new RasterizeOptions();
            options.resolution = resolution;
            options.transparency = useTransparency;
            options.backgroundBlack = false;
            options.clippingMask = false;
            options.convertSpotColors = false;
            options.includeLayers = false;
            options.padding = 0;

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
            if (rasterResult) {
                rasterResult.width = originalWidth;
                rasterResult.height = originalHeight;

                rasterResult.position = [
                    originalLeft,
                    originalTop
                ];
            }

            converted++;

            if (rasterResult) {
                outputItems.push(rasterResult);

                if (lockStateKnown && wasLocked) {
                    try {
                        rasterResult.locked = true;
                    } catch (lockError) {
                        addDetail(
                            "Warning: could not restore lock state for " +
                            itemName + " [" + itemType + "]"
                        );
                    }
                }
            }
        } catch (e2) {
            failed++;

            var errorText = "Unknown error";
            try {
                errorText = e2.message || e2.toString();
            } catch (e3) {}

            addDetail(
                "Failed: " + itemName + " [" + itemType + "] - " +
                errorText
            );

            if (lockStateKnown && wasLocked) {
                try {
                    item.locked = true;
                } catch (e4) {}
            }
        }
    }

    var result = "Rasterization Complete\n\n";
    result += "Resolution: " + resolution + " DPI\n";
    result += "Background: " +
        (useTransparency ? "Transparent" : "White") + "\n";
    result += "Selected items: " + targets.length + "\n";
    result += "Successfully rasterized: " + converted + "\n";
    result += "Skipped: " + skipped + "\n";
    result += "Failed: " + failed + "\n";

    if (details.length > 0) {
        result += "\nDetails:\n";
        for (var d = 0; d < details.length; d++) {
            result += "- " + details[d] + "\n";
        }

        if (details.length >= 30) {
            result += "Additional details were omitted.\n";
        }
    }

    result += "\nThis Script Made By Ibraheem Sabbah - www.linkedin.com/in/ibraheem-sabbah";

    alert(result);

})();
