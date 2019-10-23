#target indesign 
//#target InDesign-10 
#targetengine "imcompose"

/* 
	Copies images from one document into a new document.
	Made for Adobe InDesign CC 2019 on Windows 10.
	Copyright (C) 2019, Maximilian Bauer
	
	This program is free software: you can redistribute it and/or modify
	it under the terms of the GNU General Public License as published by
	the Free Software Foundation, either version 3 of the License, or
	(at your option) any later version.
	
	This program is distributed in the hope that it will be useful,
	but WITHOUT ANY WARRANTY; without even the implied warranty of
	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	GNU General Public License for more details.
	
	You should have received a copy of the GNU General Public License
	along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

//Switch to millimeters as unit for this script.
app.scriptPreferences.measurementUnit = MeasurementUnits.MILLIMETERS;

//The localisations of all strings used by this script.
var gui =
{
	menuItem : {"en": "Image &collection...", "de": "Bilder&sammlung..."},
	alternativeMenuItem : {"en": "New picture &collection", "de": "Neue Bilder&sammlung"},
	windowTitle : {"en": "Create a image collection from an existing InDesign document", "de": "Bildersammlung aus existierendem InDesign-Dokument erstellen"},
	pagesPanelTitle : {"en": "Pages", "de": "Seiten" },
	pagesComboBoxElementAll : {"en": "All pages", "de": "Alle Seiten"},
	pagesComboBoxElementCurrent : {"en": "Current page", "de": "Aktuelle Seite"},
	pagesComboBoxElementManual : {"en": "Pages", "de": "Seiten"},
	optionsPanelTitle : {"en": "Options", "de": "Optionen"},
	optionsLabelBleed : {"en": "Bleed:", "de": "Anschnitt:"},
	optionsCheckboxKeepOriginalImageSize : {"en": "Use image size of document", "de": "Bildgröße aus Originaldokument verwenden"},
	optionsLabelIgnoreBelowWidth : {"en": "Ignore images with a width of less than ", "de": "Ignoriere Bilder mit einer Breite unter "},
	optionsLabelIgnoreBelowHeight: {"en": "Ignore images with a height of less than ", "de": "Ignoriere Bilder mit einer Höhe unter "},
	optionsCheckboxIgnoreImagesWithEffects: {"en": "Ignore images with effects", "de": "Bilder mit Effekten ignorieren"},
	buttonExecute: {"en": "Create document", "de": "Dokument erstellen"},
	buttonCancel: {"en": "Cancel", "de": "Abbrechen"},
	invalidManualPagesMessage: {"en": "The specified page names/numbers are invalid or missing. Please mind that page ranges are not supported.",
		"de": "Die angegebenen Seitenbezeichnungen sind ungültig oder fehlen. Bitte beachten Sie, dass Bereichsangaben nicht unterstützt werden."},
	invalidManualPagesTitle: {"en": "Invalid page names", "de": "Ungültige Seitenbezeichnungen"},
	invalidValuesMessage: {"en": "One or more of the entered unit values are invalid or missing. "
		+ "Please ensure all fields are only filled with numeric values using the '.' as decimal mark and try again.", 
		"de": "Die eingegebenen Maße sind ungültig oder fehlen. Bitte überprüfen Sie, "
		+ "dass alle Maße korrekt als numerische Werte angegeben sind, '.' statt ',' als Dezimaltrennzeichen verwendet wurde und versuchen es erneut."},
	invalidValuesTitle: {"en": "Invalid unit values", "de": "Ungültige Maße"},
	openBaseFileDialogTitle: {"en": "Please select the InDesign file from which the images should be used", 
		"de": "Bitte wählen Sie die InDesign-Datei, aus der die Bilder verwendet werden sollen"}
}

//The main functionality of this script as function.
var doOperation = function(sourceDocument, operationParameters)
{
	if (!sourceDocument) throw "No source document specified.";
	if (!operationParameters) throw "No operation parameters specified.";

	//Create a new (hidden) document, into which the images are copied into.
	var targetDocument = app.documents.add(!operationParameters.hideNewDocumentDuringProcess);
	targetDocument.name = "IMCG";

	//Set the parameters of the document so that it has no slug/margins and only one page in the master spread.
	with (targetDocument.documentPreferences)
	{
		documentBleedUniformSize = true, documentBleedBottomOffset = operationParameters.bleed.toString(),
		documentSlugUniformSize = true, slugBottomOffset = "0mm",
		facingPages = false
	}
	with(targetDocument.marginPreferences)
	{
		top = 0, left = 0, bottom = 0, right = 0
	}

	for (var si = 1; si < targetDocument.masterSpreads.item(0).pages.length; si++)
		targetDocument.masterSpreads.item(0).pages.item(si).remove();
	 
	targetDocument.masterSpreads.item(0).pages.item(0).marginPreferences.top = 0;
	targetDocument.masterSpreads.item(0).pages.item(0).marginPreferences.left = 0;
	targetDocument.masterSpreads.item(0).pages.item(0).marginPreferences.bottom = 0;
	targetDocument.masterSpreads.item(0).pages.item(0).marginPreferences.right = 0;

	//Iterate through the spreads of the source document, searching for image elements.
	for (var s = 0; s < sourceDocument.spreads.length; s++)
	{
		var currentSpread = sourceDocument.spreads.item(s);
		
		//Calculate the dimensions of the current spread, which is required to crop the image on the spread below.
		var spreadDimensions = { top: 0, left: 0, bottom: 0, right: 0 };	

		for (var p = 0; p < currentSpread.pages.length; p++)
		{
			var currentPage = currentSpread.pages[p];
			
			spreadDimensions.top = Math.min(spreadDimensions.top, currentPage.bounds[0]);
			spreadDimensions.left = Math.min(spreadDimensions.left, currentPage.bounds[1]);
			spreadDimensions.bottom = Math.max(spreadDimensions.bottom, currentPage.bounds[2]);
			spreadDimensions.right = Math.max(spreadDimensions.right, currentPage.bounds[3]);
		}

		//Iterate through all image objects on the current spread.
		for (var i = 0; i < currentSpread.allPageItems.length; i++)
		{
			var currentItem = currentSpread.allPageItems[i];
					
			//Calculate how much of the current image is in the bleed area of the original document and needs to be cropped away.
			var bleedCrop =
			{
				top : currentItem.geometricBounds[0] - Math.max(currentItem.geometricBounds[0], 0),
				left : currentItem.geometricBounds[1] - Math.max(currentItem.geometricBounds[1], 0),
				bottom : currentItem.geometricBounds[2] - Math.min(currentItem.geometricBounds[2], spreadDimensions.bottom),
				right : currentItem.geometricBounds[3] - Math.min(currentItem.geometricBounds[3], spreadDimensions.right)
			}
		
			//Calculate the correct geometric bounds of the pasted image and its width and height.
			var itemWidth = currentItem.geometricBounds[3] - currentItem.geometricBounds[1] + bleedCrop.left - bleedCrop.right;
			var itemHeight = currentItem.geometricBounds[2] - currentItem.geometricBounds[0] + bleedCrop.top - bleedCrop.bottom;
			
			
			//If the selected item doesn't contain an image, it's skipped.
			if (currentItem.allGraphics.length == 0) continue;
			
			//If the image size in the original document is smaller than ignoreBelowWidth/ignoreBelowHeight, it's skipped.
			if (itemWidth < operationParameters.ignoreBelowWidth ||
				itemHeight < operationParameters.ignoreBelowHeight)
				continue;
			
			//If one or more page names are defined, only the images from these pages are be copied.
			if (operationParameters.pageNames.length > 0)
			{
				var copyImage = false;
				for (var n = 0; n < operationParameters.pageNames.length; n++)
					copyImage = copyImage || (operationParameters.pageNames[n] == currentItem.parentPage.name);
				if (!copyImage) continue;
			}			

			
			//Get the first empty page in the target document (or create one).
			var targetPage = null;
			for (var tp = 0; tp < targetDocument.pages.length; tp++)
			{
				var targetPageCandidate = targetDocument.pages.item(tp);
				if (targetPageCandidate.pageItems.length == 0)
				{
					targetPage = targetPageCandidate;
					break;
				}
			}
			if (!targetPage) targetPage = targetDocument.pages.add();

			//Prepare the page so that the image will fit on it and there's no white spaces.
			targetPage.adjustLayout(
			{ 
				bottomMargin : 0, leftMargin : 0, topMargin : 0, rightMargin : 0, 
				width : itemWidth + 'mm', height : itemHeight + 'mm'
			});
			
			//Copy the current image to the new target page, correct the bounds and move it to the top-left.
			var duplicatedItem = currentItem.duplicate(targetPage);
			duplicatedItem.frameFittingOptions.autoFit = false;
			duplicatedItem.move([bleedCrop.left, bleedCrop.top]);
			duplicatedItem.geometricBounds =
			[
				duplicatedItem.geometricBounds[0] - bleedCrop.top,
				duplicatedItem.geometricBounds[1] - bleedCrop.left,
				duplicatedItem.geometricBounds[2] - bleedCrop.bottom,
				duplicatedItem.geometricBounds[3] - bleedCrop.right
			];
			
			//If the parameter is set, scale the page and the contained image so that the image is scaled to 100%.
			if (!operationParameters.keepOriginalImageSize)
			{
				var largestItemScaleFactor = Math.max(duplicatedItem.allGraphics[0].horizontalScale,
					duplicatedItem.allGraphics[0].verticalScale);
				var pageScalingFactor = 100.0/largestItemScaleFactor;

				duplicatedItem.frameFittingOptions.autoFit = true;
				targetPage.resize(CoordinateSpaces.INNER_COORDINATES, AnchorPoint.CENTER_ANCHOR,
					ResizeMethods.MULTIPLYING_CURRENT_DIMENSIONS_BY, [pageScalingFactor, pageScalingFactor]);
				duplicatedItem.resize(CoordinateSpaces.INNER_COORDINATES, AnchorPoint.CENTER_ANCHOR,
					ResizeMethods.MULTIPLYING_CURRENT_DIMENSIONS_BY, [pageScalingFactor, pageScalingFactor]);
				duplicatedItem.frameFittingOptions.autoFit = false;
			}
			
			//Extend the borders of the image item by the defined bleed.
			//This must be done here or the bleed of the item could get scaled.
			duplicatedItem.geometricBounds =
			[
				duplicatedItem.geometricBounds[0] - operationParameters.bleed.value,
				duplicatedItem.geometricBounds[1] - operationParameters.bleed.value,
				duplicatedItem.geometricBounds[2] + operationParameters.bleed.value,
				duplicatedItem.geometricBounds[3] + operationParameters.bleed.value
			];
		}
	}
	
	if (operationParameters.hideNewDocumentDuringProcess) targetDocument.windows.add();
}

//Creates the GUI window, which enables the user to configure how the script is executed.
var showCreateDocumentFromImagesConfigurationDialog = function(sourceDocument)
{
	var createUnitDropdown = function(parentElement)
	{
		var unitDropdown = parentElement.add("dropdownlist");
		unitDropdown.add ("item", "mm");
		unitDropdown.add ("item", "cm");
		unitDropdown.add ("item", "in");
		unitDropdown.add ("item", "pt");    
		unitDropdown.selection = 0;
		return unitDropdown;
	}

	var createUnitInputGroup = function(parentElement, label, inputLength)
	{
		var unitInputGroup = parentElement.add("group");
		unitInputGroup.orientation = "row";
		unitInputGroup.label = unitInputGroup.add("statictext", undefined, label);
		unitInputGroup.field = unitInputGroup.add("edittext");
		unitInputGroup.field.characters = inputLength;
		unitInputGroup.unit = createUnitDropdown(unitInputGroup);
		return unitInputGroup;
	}
	
	//Initialize the window with all elements
	var window = new Window("palette", localize(gui.windowTitle));
	window.orientation = "column";
	window.alignChildren = "right";
	
	window.panel1 = window.add("panel", undefined, localize(gui.pagesPanelTitle));
	window.panel1.orientation = "column";
	window.panel1.alignChildren = "left";
	window.panel1.minimumSize.width = 450;
	window.panel1.allPagesRadioButton = window.panel1.add("radiobutton", undefined, localize(gui.pagesComboBoxElementAll));
	window.panel1.allPagesRadioButton.value = true;
	window.panel1.allPagesRadioButton.onClick = function() 
	{
		window.panel1.currentPageRadioButton.value = false;
		window.panel1.manualPagesGroup.manualPagesRadioButton.value = false;
		window.panel1.manualPagesGroup.pagesField.enabled = false;
	};
	window.panel1.currentPageRadioButton = window.panel1.add("radiobutton", undefined, localize(gui.pagesComboBoxElementCurrent));
	window.panel1.currentPageRadioButton.onClick = function() 
	{
		window.panel1.allPagesRadioButton.value = false;
		window.panel1.manualPagesGroup.manualPagesRadioButton.value = false;
		window.panel1.manualPagesGroup.pagesField.enabled = false;
	};
	window.panel1.manualPagesGroup = window.panel1.add("group");
	window.panel1.manualPagesGroup.orientation = "row";
	window.panel1.manualPagesGroup.manualPagesRadioButton = window.panel1.manualPagesGroup.add("radiobutton", undefined, localize(gui.pagesComboBoxElementManual));
	window.panel1.manualPagesGroup.manualPagesRadioButton.onClick = function() 
	{
		window.panel1.allPagesRadioButton.value = false;
		window.panel1.currentPageRadioButton.value = false;
		window.panel1.manualPagesGroup.pagesField.enabled = true;
	};
	window.panel1.manualPagesGroup.pagesField = window.panel1.manualPagesGroup.add("edittext");
	window.panel1.manualPagesGroup.pagesField.characters = 25;
	window.panel1.manualPagesGroup.pagesField.enabled = false;
	
	window.panel2 = window.add("panel", undefined, localize(gui.optionsPanelTitle));
	window.panel2.orientation = "column";
	window.panel2.alignChildren = "left";
	window.panel2.minimumSize.width = 450;
	window.panel2.bleedGroup = createUnitInputGroup(window.panel2, localize(gui.optionsLabelBleed), 3);
	window.panel2.bleedGroup.field.text = "0";
	window.panel2.ignoreBelowWidthGroup = createUnitInputGroup(window.panel2, localize(gui.optionsLabelIgnoreBelowWidth), 4);
	window.panel2.ignoreBelowWidthGroup.field.text = "20";
	window.panel2.ignoreBelowHeightGroup = createUnitInputGroup(window.panel2, localize(gui.optionsLabelIgnoreBelowHeight), 4);
	window.panel2.ignoreBelowHeightGroup.field.text = "20";
	window.panel2.ignoreWithEffectsCheckBox = window.panel2.add("checkbox", undefined, localize(gui.optionsCheckboxIgnoreImagesWithEffects));
	window.panel2.ignoreWithEffectsCheckBox.value = false;
	window.panel2.ignoreWithEffectsCheckBox.enabled = false;
	window.panel2.keepOriginalImageSizeCheckBox = window.panel2.add("checkbox", undefined, localize(gui.optionsCheckboxKeepOriginalImageSize));
	
	window.buttonGroup = window.add("group");
	window.buttonGroup.alignChildren = "right";
	window.buttonGroup.orientation = "row";
	window.buttonGroup.cancelButton = window.buttonGroup.add("button", undefined, localize(gui.buttonCancel));
	window.buttonGroup.executeButton = window.buttonGroup.add("button", undefined, localize(gui.buttonExecute));
	window.buttonGroup.executeButton.onClick = function() 
	{
		var selectedPageNames = [];
		if (window.panel1.manualPagesGroup.manualPagesRadioButton.value)
		{
			if (!window.panel1.manualPagesGroup.pagesField.text.match(new RegExp("[0-9,A-Za-z]+")))
			{
				alert(localize(gui.invalidManualPagesMessage), localize(gui.invalidManualPagesTitle), false);
				return;
			} 
			else
			{
				selectedPageNames = window.panel1.manualPagesGroup.pagesField.text.split(',');
				for (var pn = 0; pn < selectedPageNames.length; pn++)
					selectedPageNames[pn] = selectedPageNames[pn].replace(' ', '');
			}
		}
		else if (window.panel1.currentPageRadioButton.value)
		{
			selectedPageNames = [app.activeWindow.activePage.name];
		}
		
		var operationParametersUser;
		try
		{
			operationParametersUser =
			{
				bleed : new UnitValue(new Number(window.panel2.bleedGroup.field.text), window.panel2.bleedGroup.unit.selection.text),
				keepOriginalImageSize: window.panel2.keepOriginalImageSizeCheckBox.value,
				ignoreBelowWidth : new UnitValue(new Number(window.panel2.ignoreBelowWidthGroup.field.text), window.panel2.ignoreBelowWidthGroup.unit.selection.text),
				ignoreBelowHeight : new UnitValue(new Number(window.panel2.ignoreBelowHeightGroup.field.text), window.panel2.ignoreBelowHeightGroup.unit.selection.text),
				ignoreWithEffects: window.panel2.ignoreWithEffectsCheckBox.value,
				hideNewDocumentDuringProcess : true,
				pageNames : selectedPageNames
			};
		
			if (operationParametersUser.bleed.value !== operationParametersUser.bleed.value) throw "bleed";
			if (operationParametersUser.ignoreBelowWidth.value !== operationParametersUser.ignoreBelowWidth.value) throw "ignoreBelowWidth.";
			if (operationParametersUser.ignoreBelowHeight.value !== operationParametersUser.ignoreBelowHeight.value) throw "ignoreBelowHeight";
		} 
		catch (error)
		{
			alert(localize(gui.invalidValuesMessage)+ " (" + error + ")", localize(gui.invalidValuesTitle), false);
			return;
		}
	
		window.enabled = false;

		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
		try { doOperation(sourceDocument, operationParametersUser); }
		catch (error)
		{
			alert(error, "Operation failed", true);
		}
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;  
		
		window.close();
	};
	window.buttonGroup.cancelButton.onClick = function()
	{
		window.close();
	};
	
	window.show();
}

//The "main function", which checks if a document is opened (or opens one) and then opens the main dialog window.
var main = function()
{
	var currentDocument = null;

	try { currentDocument = app.activeDocument; }
	catch (error)
	{
		var documentFile = File.openDialog(localize(gui.openBaseFileDialogTitle), "InDesign:*.indp;*.idap;*.indd;*.indt;*.indl");
		try 
		{ 
			if (documentFile) 
			{
				app.open(documentFile);
				currentDocument = app.activeDocument;
			}
		}
		catch (error)
		{
			alert("Error.");
		}
	}

	if (currentDocument)
	{
		showCreateDocumentFromImagesConfigurationDialog(currentDocument);
	}
}

//Depending on whether the script is run as startup script or over the script panel, a new menu item is added.
var isRunAtStartup = false;
try 
{ 
	var parentFolderName = app.activeScript.parent.name.toLowerCase().replace("%20", " ");
	isRunAtStartup = (parentFolderName == "startup scripts");
}
catch (error) { isRunAtStartup = false; }

if (isRunAtStartup)
{

	
	var mainMenuBar = app.menus.item("$ID/Main");
	var fileSubmenu = mainMenuBar.submenus.item("$ID/" + app.translateKeyString("$ID/&File"));
	var newFileSubmenu = fileSubmenu.submenus.item("$ID/" + app.translateKeyString("$ID/&New"));
	if (newFileSubmenu) 
	{
		var newDocumentFromImagesAction = app.scriptMenuActions.add(localize(gui.menuItem));
		newDocumentFromImagesAction.eventListeners.add("onInvoke", main);
		newFileSubmenu.menuItems.add(newDocumentFromImagesAction);
	}
	else 
	{
		var newDocumentFromImagesAction = app.scriptMenuActions.add(localize(gui.alternativeMenuItem));
		newDocumentFromImagesAction.eventListeners.add("onInvoke", main);
		fileSubmenu.menuItems.add(newDocumentFromImagesAction, LocationOptions.after, fileSubmenu.menuItems.item(0));
	}
}
else
{
	main();
}