// ============================================
// CONFIGURATION: Set your OpenAI API key here
// ============================================
// Replace 'YOUR_API_KEY_HERE' with your actual OpenAI API key before use.
// Do NOT commit real keys to source control.
const OPENAI_API_KEY = "YOUR_API_KEY_HERE";

function onOpen() {
  DocumentApp.getUi()
    .createMenu("Resume Assistant")
    .addItem("Open Assistant", "showSidebar")
    .addItem("Authorize Script", "authorizeScript")
    .addToUi();
}

// This function triggers authorization - MUST be run from Apps Script editor first time
function authorizeScript() {
  try {
    // Try to access PropertiesService to trigger authorization
    const props = PropertiesService.getScriptProperties();
    const testKey = props.getProperty("OPENAI_KEY");
    
    if (testKey) {
      DocumentApp.getUi().alert("✅ Script is authorized! API key found and ready to use.");
    } else {
      DocumentApp.getUi().alert("⚠️ Script is authorized, but no API key found.\n\nYou can set it using 'Configure API Key' menu option, or manually in Apps Script:\n1. Go to Project Settings (gear icon)\n2. Script Properties\n3. Add OPENAI_KEY");
    }
  } catch (error) {
    DocumentApp.getUi().alert("❌ Authorization needed!\n\nPlease:\n1. Go to Extensions > Apps Script\n2. Select 'testAuthorization' from function dropdown\n3. Click Run (▶️)\n4. Click 'Review Permissions' and authorize");
  }
}

// Simple test function to trigger authorization - run this from Apps Script editor
function testAuthorization() {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty("OPENAI_KEY");
  Logger.log("API Key found: " + (key ? "Yes (starts with " + key.substring(0, 7) + "...)" : "No"));
  return "Authorization successful! API key " + (key ? "found" : "not found");
}

// Alternative authorization function - run this if testAuthorization doesn't work
function authorizeAndCheck() {
  try {
    // This will trigger authorization if needed
    const props = PropertiesService.getScriptProperties();
    const key = props.getProperty("OPENAI_KEY");
    
    if (key) {
      Logger.log("✅ SUCCESS! Authorization works. API key found.");
      return "✅ Authorization successful! API key found and ready to use.";
    } else {
      Logger.log("⚠️ Authorization works, but no API key found.");
      return "⚠️ Authorization works, but OPENAI_KEY not found in Script Properties.";
    }
  } catch (error) {
    Logger.log("❌ ERROR: " + error.toString());
    // Re-throw to trigger authorization dialog
    throw new Error("Authorization needed. Error: " + error.toString());
  }
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("sidebar")
    .setTitle("Resume Assistant");
  DocumentApp.getUi().showSidebar(html);
}

// Helper function to get API key
function getApiKey() {
  // 1) Hardcoded constant (replace before use)
  if (OPENAI_API_KEY && OPENAI_API_KEY !== "YOUR_API_KEY_HERE") return OPENAI_API_KEY;
  // 2) Script/Document properties (editor context)
  try {
    const docKey = PropertiesService.getDocumentProperties().getProperty("OPENAI_KEY");
    if (docKey) return docKey;
    const scriptKey = PropertiesService.getScriptProperties().getProperty("OPENAI_KEY");
    if (scriptKey) return scriptKey;
  } catch (e) {
    Logger.log("PropertiesService not accessible: " + e.toString());
  }
  return null;
}

// Helper function to save API key to PropertiesService (editor context)
function setApiKey(key) {
  Logger.log("setApiKey called");
  if (!key || key.trim() === "") {
    throw new Error("API key cannot be empty");
  }
  const trimmed = key.trim();
  try {
    PropertiesService.getDocumentProperties().setProperty("OPENAI_KEY", trimmed);
    PropertiesService.getScriptProperties().setProperty("OPENAI_KEY", trimmed);
    Logger.log("Saved API key to PropertiesService");
    return true;
  } catch (e) {
    Logger.log("Could not save to PropertiesService: " + e.toString());
    throw new Error("Failed to save API key to PropertiesService: " + e.toString());
  }
}

// Diagnostic function - call this from sidebar to debug
function diagnosticCheck() {
  const results = {
    step: "Starting diagnostic",
    errors: [],
    successes: []
  };
  
  try {
    results.step = "Checking Document Custom Properties";
    const doc = DocumentApp.getActiveDocument();
    const docKey = doc.getCustomProperty("OPENAI_KEY");
    if (docKey) {
      results.successes.push("Document Custom Property: API key found");
    } else {
      results.errors.push("Document Custom Property: No API key found");
    }
  } catch (e) {
    results.errors.push("Document Custom Property access failed: " + e.toString());
  }
  
  try {
    results.step = "Checking Document Properties";
    const docProps = PropertiesService.getDocumentProperties();
    const docKey = docProps.getProperty("OPENAI_KEY");
    if (docKey) {
      results.successes.push("Document Properties: API key found");
    } else {
      results.errors.push("Document Properties: No API key found");
    }
  } catch (e) {
    results.errors.push("Document Properties access failed: " + e.toString());
  }
  
  try {
    results.step = "Checking Script Properties";
    const scriptProps = PropertiesService.getScriptProperties();
    const scriptKey = scriptProps.getProperty("OPENAI_KEY");
    if (scriptKey) {
      results.successes.push("Script Properties: API key found");
    } else {
      results.errors.push("Script Properties: No API key found");
    }
  } catch (e) {
    results.errors.push("Script Properties access failed: " + e.toString());
  }
  
  try {
    results.step = "Checking Document access";
    const doc = DocumentApp.getActiveDocument();
    const title = doc.getName();
    results.successes.push("Document access: OK (title: " + title + ")");
  } catch (e) {
    results.errors.push("Document access failed: " + e.toString());
  }
  
  // Log everything
  Logger.log("=== DIAGNOSTIC RESULTS ===");
  Logger.log("Step: " + results.step);
  Logger.log("Successes: " + JSON.stringify(results.successes));
  Logger.log("Errors: " + JSON.stringify(results.errors));
  Logger.log("========================");
  
  return results;
}

function processJobDescription(jdText) {
  Logger.log("=== processJobDescription START ===");
  Logger.log("JD length: " + (jdText ? jdText.length : 0));
  
  try {
    Logger.log("Step 1: Getting API key");
    const apiKey = getApiKey();
    Logger.log("API key result: " + (apiKey ? "Found (length: " + apiKey.length + ")" : "Not found"));
    
    if (!apiKey) {
      Logger.log("ERROR: No API key found");
      throw new Error("OpenAI API key not found. Please set OPENAI_API_KEY constant at the top of Code.gs");
    }
    
    Logger.log("Step 2: Accessing document");
    let doc;
    try {
      doc = DocumentApp.getActiveDocument();
      Logger.log("Document access: OK - " + doc.getName());
    } catch (docError) {
      Logger.log("Document access ERROR: " + docError.toString());
      throw new Error("Cannot access document: " + docError.toString());
    }
    const body = doc.getBody();
    const resumeText = body.getText();

    if (!jdText || jdText.trim() === "") {
      throw new Error("Please provide a job description");
    }

    const prompt = `You are a resume-editing assistant. 
Given the base resume and the job description, provide specific, actionable improvements.

IMPORTANT: Return your response as a JSON array of suggestions. Each suggestion should have:
- "original": the exact text from the resume to replace (copy it exactly as it appears)
- "suggestion": the improved version
- "reason": brief explanation of why this change helps

Format:
[
  {
    "original": "exact text from resume",
    "suggestion": "improved version",
    "reason": "why this helps"
  }
]

BASE RESUME:
${resumeText}

JOB DESCRIPTION:
${jdText}

Return ONLY the JSON array, no other text.`;

    Logger.log("Step 4: Calling OpenAI");
    const response = callOpenAI(prompt);
    Logger.log("Step 5: Parsing suggestions");
    const suggestions = parseSuggestions(response);
    Logger.log("Step 6: Applying suggestions");
    applySuggestionsToDoc(suggestions);
    Logger.log("SUCCESS: Applied " + suggestions.length + " suggestions");
    return { success: true, count: suggestions.length };
  } catch (error) {
    Logger.log("ERROR in processJobDescription: " + error.toString());
    Logger.log("Error stack: " + (error.stack || "No stack trace"));
    throw error;
  }
}

function improveSelection() {
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      throw new Error("OpenAI API key not found. Please use 'Resume Assistant > Configure API Key' to set it up.");
    }
    
    const doc = DocumentApp.getActiveDocument();
    const selection = doc.getSelection();
    
    if (!selection || selection.getRangeElements().length === 0) {
      throw new Error("Please select some text first");
    }

    const rangeElements = selection.getRangeElements();
    let selectedText = "";
    const textElements = [];
    
    rangeElements.forEach(el => {
      const element = el.getElement();
      if (element.getType() === DocumentApp.ElementType.TEXT) {
        const textElement = element.asText();
        const start = el.getStartOffset();
        const end = el.getEndOffsetInclusive();
        selectedText += textElement.getText().substring(start, end + 1);
        textElements.push({
          element: textElement,
          start: start,
          end: end
        });
      }
    });

    if (!selectedText || selectedText.trim() === "") {
      throw new Error("Selected text is empty");
    }

    const prompt = `Rewrite this resume bullet to be stronger, concise, and achievement-driven.
Focus on quantifiable achievements and action verbs.
Return ONLY the improved version, no explanations.

"${selectedText}"`;

    const newText = callOpenAI(prompt);
    if (!newText || newText.trim() === "") {
      throw new Error("No response from AI");
    }
    
    replaceSelectionWithFormatting(textElements, newText.trim());
    return { success: true };
  } catch (error) {
    Logger.log("Error in improveSelection: " + error.toString());
    throw error;
  }
}

function replaceSelectionWithFormatting(textElements, newText) {
  if (textElements.length === 0) return;
  
  const firstElement = textElements[0].element;
  const lastElement = textElements[textElements.length - 1].element;
  const startOffset = textElements[0].start;
  const endOffset = textElements[textElements.length - 1].end;
  
  // Get formatting from the first character of selection
  const attributes = firstElement.getAttributes(startOffset);
  
  // Delete all selected text
  textElements.forEach(({ element, start, end }) => {
    element.deleteText(start, end);
  });
  
  // Insert new text with preserved formatting
  firstElement.insertText(startOffset, newText);
  
  // Apply formatting to the new text
  const newEnd = startOffset + newText.length - 1;
  Object.keys(attributes).forEach(key => {
    if (attributes[key] !== null) {
      firstElement.setAttributes(startOffset, newEnd, { [key]: attributes[key] });
    }
  });
}

function parseSuggestions(response) {
  try {
    // Try to extract JSON from response (in case there's extra text)
    const jsonMatch = response.match(/\[[\s\S]*\]/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
    return JSON.parse(response);
  } catch (error) {
    Logger.log("Error parsing suggestions: " + error.toString());
    // Fallback: try to parse as simple text suggestions
    return [];
  }
}

function applySuggestionsToDoc(suggestions) {
  if (!suggestions || suggestions.length === 0) {
    throw new Error("No suggestions to apply");
  }
  
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  let appliedCount = 0;
  
  suggestions.forEach(suggestion => {
    try {
      const original = suggestion.original;
      const replacement = suggestion.suggestion;
      
      if (!original || !replacement) return;
      
      const found = body.findText(original);
      if (found) {
        const element = found.getElement();
        if (element.getType() === DocumentApp.ElementType.TEXT) {
          const textElement = element.asText();
          const start = found.getStartOffset();
          const end = found.getEndOffsetInclusive();
          
          // Get formatting from original text
          const attributes = textElement.getAttributes(start);
          
          // Replace text
          textElement.deleteText(start, end);
          textElement.insertText(start, replacement);
          
          // Apply formatting to new text
          const newEnd = start + replacement.length - 1;
          Object.keys(attributes).forEach(key => {
            if (attributes[key] !== null) {
              textElement.setAttributes(start, newEnd, { [key]: attributes[key] });
            }
          });
          
          appliedCount++;
        }
      }
    } catch (error) {
      Logger.log("Error applying suggestion: " + error.toString());
    }
  });
  
  if (appliedCount === 0) {
    throw new Error("Could not find any matching text to replace. Make sure the original text matches exactly.");
  }
}

// Check if script is authorized - call this from sidebar first
function checkAuthorization() {
  const key = getApiKey();
  if (!key) {
    return {
      authorized: true,
      hasKey: false,
      message: "API key not found. Please set OPENAI_API_KEY constant at the top of Code.gs"
    };
  }
  return { 
    authorized: true, 
    hasKey: true,
    message: "Ready to use!"
  };
}

// This function triggers authorization when called from sidebar (no DocumentApp access)
function triggerAuthorizationFromSidebar() {
  try {
    Logger.log("triggerAuthorizationFromSidebar: Starting (sidebar-safe)");
    const key = getApiKey();
    Logger.log("API key check: " + (key ? "Found" : "Not found"));
    if (!key) {
      return {
        success: true,
        hasKey: false,
        message: "API key not found. Please set OPENAI_API_KEY constant at the top of Code.gs"
      };
    }
    return {
      success: true,
      hasKey: true,
      message: "Authorization successful! API key found (hardcoded)"
    };
  } catch (error) {
    Logger.log("triggerAuthorizationFromSidebar ERROR: " + error.toString());
    Logger.log("Error stack: " + (error.stack || "No stack"));
    throw error;
  }
}

function callOpenAI(prompt) {
  Logger.log("callOpenAI: Starting");
  
  const API_KEY = getApiKey();
  Logger.log("callOpenAI: API key check: " + (API_KEY ? "Found" : "Not found"));
  
  if (!API_KEY) {
    Logger.log("callOpenAI: ERROR - No API key found");
    throw new Error("OpenAI API key not configured. Please use 'Resume Assistant > Configure API Key' to set it up.");
  }
  
  Logger.log("callOpenAI: API key found, making request");

  const url = "https://api.openai.com/v1/chat/completions";

  const payload = {
    model: "gpt-4o-mini",  // Fixed: was "gpt-4.1-mini"
    messages: [
      { role: "system", content: "You are a resume editing assistant. Always return valid JSON when requested." },
      { role: "user", content: prompt }
    ],
    temperature: 0.2
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + API_KEY
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  if (responseCode !== 200) {
    const errorData = JSON.parse(responseText);
    throw new Error(`OpenAI API error: ${errorData.error?.message || responseText}`);
  }

  const data = JSON.parse(responseText);
  
  if (!data.choices || !data.choices[0] || !data.choices[0].message) {
    throw new Error("Unexpected response format from OpenAI API");
  }

  return data.choices[0].message.content;
}

function showApiKeyDialog() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial; padding: 20px;">
      <h3>Configure OpenAI API Key</h3>
      <p>Enter your OpenAI API key. You can get one from <a href="https://platform.openai.com/api-keys" target="_blank">OpenAI Platform</a></p>
      <input type="password" id="apiKey" style="width: 100%; padding: 8px; margin: 10px 0;" placeholder="sk-...">
      <br>
      <button onclick="saveKey()" style="padding: 8px 16px; margin-top: 10px;">Save</button>
      <button onclick="google.script.host.close()" style="padding: 8px 16px; margin-top: 10px;">Cancel</button>
      <p id="status" style="margin-top: 10px; color: green;"></p>
    </div>
    <script>
      function saveKey() {
        const key = document.getElementById("apiKey").value.trim();
        if (!key) {
          alert("Please enter an API key");
          return;
        }
        google.script.run
          .withSuccessHandler(function(result) {
            const message = result && result.message ? result.message : "API key saved successfully!";
            document.getElementById("status").textContent = message;
            document.getElementById("status").style.color = "green";
            setTimeout(() => google.script.host.close(), 2000);
          })
          .withFailureHandler(function(error) {
            console.error("Save error:", error);
            const errorMsg = error.message || error.toString() || "Unknown error occurred";
            document.getElementById("status").textContent = "Error: " + errorMsg;
            document.getElementById("status").style.color = "red";
            alert("Failed to save API key:\n" + errorMsg + "\n\nCheck Apps Script execution log for details.");
          })
          .saveApiKey(key);
      }
    </script>
  `)
    .setWidth(400)
    .setHeight(250);
  
  DocumentApp.getUi().showModalDialog(html, "API Key Configuration");
}

function saveApiKey(key) {
  Logger.log("saveApiKey called");
  try {
    if (!key || key.trim() === "") {
      throw new Error("API key cannot be empty");
    }
    
    // Use the helper function which saves to footer and PropertiesService
    setApiKey(key.trim());
    Logger.log("saveApiKey completed successfully");
    return { success: true, message: "API key saved successfully!" };
  } catch (error) {
    Logger.log("ERROR in saveApiKey: " + error.toString());
    throw error;
  }
}