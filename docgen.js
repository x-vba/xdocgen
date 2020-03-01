"use strict";

// Version 1.2.0

/**
 * This function takes VBA source code and converts the relevant features and
 * the XDocGen tags into a JSON format to allow automatic documentation
 * generation.
 *
 * @author Anthony Mancini
 * @version 1.0.0
 * @license MIT
 * @todo Add module level XDocGen tag generation
 * @todo Add support for Module level variables, probably need anothe function for this, but
 * should be fairly easy to implement
 * @param {string} vbaSourceCode is the source code from a VBA module
 * @returns {string} a JSON string with relevant features of the procedures and
 * XDocGen tags
 */
function vbaDocGen(vbaSourceCode) {
	
	// Getting an array of Function and Sub source code
	let procedureCodeArray = getVbaFunctionAndSubProcedures(vbaSourceCode);
	
	// Generating partial XDocGen Documentation JSON for each Function and Sub, 
	// and associating them with Argument Documentation Objects which will be
	// used to fill in parameter detials
	procedureCodeArray = procedureCodeArray.map(procedureCode => {
		return [
			generateVbaProcedureDocumentationObject(procedureCode),
			generateVbaArgumentDocumentationObject(procedureCode),
		]
	});
	
	// Adjusting the XDocGen Documentation Objects to include details about
	// the parameters
	procedureCodeArray = procedureCodeArray.map(prodObj => adjustDocumentationParameters(prodObj[0], prodObj[1]));
	
	// Combining the Module Level XDocGen tag information into a single,
	// completed XDocGen Documentation Object
	let XDocGenDocumentationObject = {
		Module: generateVbaModuleDocumentationObject(vbaSourceCode),
		Procedures: procedureCodeArray,
	};
	
	// Returning a JSON of the XDocGen Documentation Object
	return JSON.stringify(XDocGenDocumentationObject, null, 2)
}


/**
 * This function takes a VBA source code string and returns an Array of strings
 * of the Procedures in the source code. Currently this function supports
 * Functions, Subs, and Properties
 * 
 * @author Anthony Mancini
 * @version 1.1.0
 * @license MIT
 * @todo Add support for other procedures like Operators
 * @param {string} vbaSourceCode is the source code from a VBA module
 * @returns {Array} an array of strings containing individual procedures (such
 * as Functions and Subs)
 */
function getVbaFunctionAndSubProcedures(vbaSourceCode) {
	// Creating a Regex to match all Functions and Subroutines
	let procedureRegex = /((Public|Private|Friend)\s){0,1}(Static\s){0,1}(Function|Sub|Property\sGet|Property\sLet|Property\sSet)\s{0,1}[a-zA-Z0-9_]*?\s{0,1}\([\S\s]*?End\s(Function|Sub|Property)/gmi
	
	return vbaSourceCode.match(procedureRegex)
}


/**
 * This function takes a single procedure source code and generates a partially
 * complete documentation Object with relevant details from the XDocGen tags.
 * 
 * @author Anthony Mancini
 * @version 1.1.1
 * @license MIT
 * @param {string} vbaProcedureCode is the source code of a single procedure
 * @returns {Object} a partially completed documentation Object. It's partially
 * complete in that it does not yet contain additional information about the
 * parameters
 */
function generateVbaProcedureDocumentationObject(vbaProcedureCode) {
	let documentationObject = {};
	
	// Getting the function details
	let procedureDetailsArray = vbaProcedureCode.split("(")[0].trim();
	
	// Getting the name of the Procedure
	let procedureName = procedureDetailsArray.split(" ")[procedureDetailsArray.split(" ").length - 1];
	documentationObject["Name"] = procedureName;
	
	procedureDetailsArray = procedureDetailsArray.split(" ").map(produceDetail => produceDetail.toLowerCase());
	
	// Getting the Scope of the Procedure
	let scopePart;
	
	if (procedureDetailsArray.includes("public")) {
		scopePart = "Public";
	} else if (procedureDetailsArray.includes("private")) {
		scopePart = "Private";
	} else if (procedureDetailsArray.includes("friend")) {
		scopePart = "Friend";
	} else {
		scopePart = "Public";
	}
	
	documentationObject["Scope"] = scopePart;
	
	
	// Determining if the Procedure is a Static procedure
	let staticPart;
	
	if (procedureDetailsArray.includes("static")) {
		staticPart = true;
	} else {
		staticPart = false;
	}
	
	documentationObject["Static"] = staticPart;
	
	
	// Determining if the Procedure is a Function or a Subroutine
	let procedurePart;
	
	if (procedureDetailsArray.includes("function")) {
		procedurePart = "Function";
	} else if (procedureDetailsArray.includes("sub")) {
		procedurePart = "Sub";
	} else if (procedureDetailsArray.includes("property")) {
		procedurePart = "Property";
	}
	
	documentationObject["Procedure"] = procedurePart;
	
	// For Properties, determining if they are Getters, Letters, or Setters
	if (procedurePart === "Property") {
		let propertyType = procedureDetailsArray[1].substr(0, 1).toUpperCase() + procedureDetailsArray[1].substr(1, procedureDetailsArray[1].length);
		documentationObject["Property"] = propertyType;
		
		// Modifying the name in cases of Properties so that the Get and Let don't
		// overwrite each other in the documentation object
		documentationObject["Name"] = `${documentationObject["Name"]}(${propertyType})`
	}
	
	// Getting the return type of the function
	let typePart = vbaProcedureCode.split(")")[1];
	
	// Could use a bit of refactoring, as a little complicated now, but handles
	// cases where the return type is on a seperate line as the As, such as As _
	if (typePart.split("\n")[0].includes(" _")) {
		typePart = typePart.split("\n")[1].trim();
	} else if (typePart.split("\n")[0].toLowerCase().includes("as ")) {
		typePart = typePart.split(/As\s/gmi)[1].trim();
	} else {
		typePart = "Variant";
	}
	
	if (typePart.includes("\n")) {
		typePart = typePart.split("\n")[0].trim();
	}
	
	if (typePart.toLowerCase().includes("as ")) {
		typePart = typePart.split(/As\s/gmi)[1].trim();
	}
	
	documentationObject["Type"] = typePart;
	
	
	// Adding the procedure source code
	documentationObject["Source"] = vbaProcedureCode;
	

	return Object.assign(documentationObject, generateXDocGenTagsObject(vbaProcedureCode)) 
}


/**
 * This function takes a single procedure source code and generates an
 * Argument documentation Object which contains details about the arguments of
 * the procedure
 * 
 * @author Anthony Mancini
 * @version 1.0.0
 * @license MIT
 * @param {string} vbaProcedureCode is the source code of a single procedure
 * @returns {Object} an Argument Documentation Object which will be used to
 * fill in the remaining details of the Documentation Object
 */
function generateVbaArgumentDocumentationObject(vbaProcedureCode) {
	let documentationTagRegex = /\([\S\s]*?\)/gmi
	let documentationObject = {};
	
	let argumentArray = vbaProcedureCode.match(documentationTagRegex)[0];
	argumentArray = argumentArray.replace("(", "").replace(")", "").replace(/_/gmi, "").split(","); 
	argumentArray = argumentArray.map(argumentLine => argumentLine.trim());
	
	argumentArray.forEach(argumentLine => {
		let modifierArray = argumentLine.split(/\sAs\s/gmi)[0].split(" ");
		modifierArray = modifierArray.map(modifierName => modifierName.toLowerCase());
		
		let optionalPart;
		if (modifierArray.includes("optional"))
			optionalPart = true;
		else
			optionalPart = false;
				
		let passingPart;
		if (modifierArray.includes("byval"))
			passingPart = "ByVal";
		else
			passingPart = "ByRef";
		
		let paramArrayPart;
		if (modifierArray.includes("paramarray"))
			paramArrayPart = true;
		else
			paramArrayPart = false;
		
		let namePart = argumentLine.split(/\sAs\s/gmi)[0].split(" ");
		namePart = namePart[namePart.length - 1].replace().trim();
		
		let arrayPart;
		if (namePart.includes("("))
			arrayPart = true;
		else
			arrayPart = false;
		
		namePart = namePart.split("(").join("").split(")").join("");
		
		let typePart;
		if (argumentLine.toLowerCase().includes(" as "))
			typePart = argumentLine.split(/\sAs\s/gmi)[1].trim();
		else
			typePart = "Variant";

		let defaultValuePart;
		if (argumentLine.toLowerCase().includes("="))
			defaultValuePart = argumentLine.split("=")[1].trim(); 
		else
			defaultValuePart = null;
		
		
		documentationObject[namePart] = {
			Name: namePart,
			Optional: optionalPart,
			Passing: passingPart,
			ParamArray: paramArrayPart,
			Type: typePart,
			Array: arrayPart,
			Default: defaultValuePart,
		};
		
	});
	
	return documentationObject
}


/**
 * This function takes a partially completed Documentation Object and an
 * Argument Documentation Object and fills in the Param details of the 
 * Documentation Object
 * 
 * @author Anthony Mancini
 * @version 1.0.0
 * @license MIT
 * @todo potentially rework this function a bit as it is currently quite nested
 * since it is handling a lot of different cases. It's the most complicated 
 * function in this program probably
 * @param {Object} documentationObject is a partially completed Documentation Object
 * @param {Object} parameterObject is an Argument Documentation Object
 * @returns {Object} a filled in Documentation Object containing Param details 
 */
function adjustDocumentationParameters(documentationObject, parameterObject) {
	
	// Used when there are more than 1 param and thus is an array
	if (Array.isArray(documentationObject["Param"])) {
		documentationObject["Param"] = documentationObject["Param"].map(paramLine => {
			let argumentName = paramLine.split(" ")[0];
			argumentName = argumentName.split("(").join("").split(")").join("");
			let argumentDescription = paramLine.substr(paramLine.indexOf(" ") + 1, paramLine.length);
			
			try {
				return Object.assign(parameterObject[argumentName], {Description: argumentDescription})
			} catch (e) {
				throw `Argument Not Found.\n\nDoes the list contain your argument?\nIf not, check that that your @Param tags match the argument names.\n\nArgument Name: ${argumentName}\nArgument List: [${Object.keys(parameterObject)}]`
			}
		});
	
	// Used when there are no @Param tags
	} else if (documentationObject["Param"] === undefined) {
		
		// Executes when the Procedure has no actual arguments
		if ("" in parameterObject) {
			documentationObject["Param"] = null;
			
		// Executes when there are parameters but there are no @Param tags
		// documenting them
		} else {
			
			// Executes for multiple parameters
			if (Object.keys(parameterObject).length > 1) {
				let parameterArray = [];
				
				for (let paramKey in parameterObject) {
					parameterArray.push(parameterObject[paramKey]);
				}
				
				documentationObject["Param"] = parameterArray;
				
			// Executes for single parameters
			} else {
				documentationObject["Param"] = parameterObject[Object.keys(parameterObject)[0]];
			}
		}
		
	// Used when there is only a single parameter and thus not an array
	} else {
		let argumentName = documentationObject["Param"].split(" ")[0];
		argumentName = argumentName.split("(").join("").split(")").join("");
		let argumentDescription = documentationObject["Param"].substr(documentationObject["Param"].indexOf(" ") + 1, documentationObject["Param"].length);
		
		try {
			documentationObject["Param"] = Object.assign(parameterObject[argumentName], {Description: argumentDescription})
		} catch (e) {
			throw `Argument Not Found.\n\nDoes the list contain your argument?\nIf not, check that that your @Param tags match the argument names.\n\nArgument Name: ${argumentName}\nArgument List: [${Object.keys(parameterObject)}]`
		}
		
	}
	
	return documentationObject
}


/**
 * This function generates the Module Level tag documentation object
 *
 * @author Anthony Mancini
 * @version 1.0.0
 * @license MIT
 * @todo clean up the code names, and perhaps move some of this stuff to a class since I
 * am reusing a bunch of code here, and especially the regexs and the tag doc generation.
 * @param {string} vbaSourceCode is the VBA source code
 * @returns {Object} a Module Documentation Object
 */
function generateVbaModuleDocumentationObject(vbaSourceCode) {
	let procedureRegex = /((Public|Private|Friend)\s){0,1}(Static\s){0,1}(Function|Sub|Property\sGet|Property\sLet|Property\sSet)\s{0,1}[a-zA-Z0-9_]*?\s{0,1}\([\S\s]*?End\s(Function|Sub|Property)/gmi
	
	let procedureMatches = vbaSourceCode.match(procedureRegex);

	if (procedureMatches !== null) {
		procedureMatches.forEach(procedureCode => {
			vbaSourceCode = vbaSourceCode.split(procedureCode).join("");
		});
	}
	

	return generateXDocGenTagsObject(vbaSourceCode);
}


/**
 * This function generates the Module Level tag documentation object
 *
 * @author Anthony Mancini
 * @version 1.0.0
 * @license MIT
 * @todo add this to a class at some point
 * @param {string} vbaCodeFragment is the remaining VBA source code after removing
 * procedures from it
 * @returns {Object} a Module Documentation Object
 */
function generateXDocGenTagsObject(vbaCodeFragment) {
	
	let documentationTagRegex = /\'\@[a-zA-Z0-9_]*?[:][\S\s]*?\n/gmi
	
	// Generating the documentation
	let tagMatches = vbaCodeFragment.match(documentationTagRegex);
	let documentationObject = {};
	
	if (tagMatches !== null) {
		tagMatches.forEach(docLine => {
			let docTag = docLine.substr(0, docLine.indexOf(":")).replace("@", "").replace("'", "").replace(":", "").trim();
			let docValue = docLine.substr(docLine.indexOf(":") + 1, docLine.length - 1).trimLeft().replace("\n", "").replace("\r", "");
			
			if (docTag in documentationObject) {
				documentationObject[docTag].push(docValue);
			} else {
				documentationObject[docTag] = [docValue];
			}
		});
	}
	
	// Adjusting the documentation so that all arrays with length of 1 are
	// combined into a single value
	for (let docTag in documentationObject) {
		if (documentationObject[docTag].length === 1) {
			documentationObject[docTag] = documentationObject[docTag][0];
		}
	}
	
	return documentationObject;
}


module.exports = {
	vbaDocGen: vbaDocGen,
};
