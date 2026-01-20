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

    const prompt = `You are a resume-editing assistant. Return ONLY valid JSON (no prose). Optimize for ATS scannability, strong action verbs, and STAR-style clarity while preserving truthful, existing content. Do NOT invent new skills, tools, or experiences.

CRITICAL: Only suggest changes for bullets that NEED improvement. If a bullet is already strong, clear, achievement-driven, and well-written, DO NOT include it in your suggestions. Only suggest changes when there's meaningful improvement to be made (weak verbs, missing metrics, unclear impact, poor ATS alignment, etc.).

HARD RULES:
- Do not add line breaks or new bullets.
- Keep formatting consistent; preserve tense/person/voice.
- If original_char_count <= 123, keep suggested_char_count <= original_char_count.
- If original_char_count > 123, keep suggested_char_count <= original_char_count * 1.15.
- Suggestion text must be a single line (no \\n).
- Self-report character counts for original and suggestion (count all characters).
- Return an EMPTY array [] if no bullets need improvement.

EXPECTED JSON ARRAY FORMAT ONLY:
[
  {
    "original": "exact text from resume",
    "suggestion": "improved version (single line, no new bullets)",
    "reason": "why this helps (ATS/impact/clarity, truthful)",
    "original_char_count": <number>,
    "suggested_char_count": <number>
  }
]

BASE RESUME:
${resumeText}

JOB DESCRIPTION:
${jdText}`;

    Logger.log("Step 4: Calling OpenAI");
    const response = callOpenAI(prompt);
    Logger.log("Step 5: Parsing suggestions");
    
    let suggestions;
    try {
      suggestions = parseSuggestions(response);
    } catch (parseError) {
      Logger.log("Parse error details: " + parseError.toString());
      // Return error info to UI instead of throwing (allows UI to show helpful message)
      return {
        success: false,
        error: "parse",
        message: parseError.message || parseError.toString(),
        count: 0,
        suggestions: []
      };
    }
    
    Logger.log("Step 6: Validating suggestions");
    const validSuggestions = validateSuggestions(suggestions);
    
    // Step 7: Analyze match percentage and missing items
    Logger.log("Step 7: Analyzing resume match to job description");
    let matchAnalysis = null;
    try {
      matchAnalysis = analyzeResumeMatch(resumeText, jdText);
    } catch (analysisError) {
      Logger.log("Match analysis error (non-fatal): " + analysisError.toString());
      // Don't fail the whole operation if analysis fails
    }
    
    Logger.log("Step 8: Returning suggestions to client (no auto-apply)");
    // Do NOT auto-apply; client will review/approve in sidebar
    return { 
      success: true, 
      count: validSuggestions.length, 
      suggestions: validSuggestions,
      matchAnalysis: matchAnalysis
    };
  } catch (error) {
    Logger.log("ERROR in processJobDescription: " + error.toString());
    Logger.log("Error stack: " + (error.stack || "No stack trace"));
    throw error;
  }
}

// Apply only the suggestions explicitly approved from the sidebar.
function applySelectedSuggestions(selectedSuggestions) {
  Logger.log("applySelectedSuggestions: start");
  try {
    if (!selectedSuggestions || !Array.isArray(selectedSuggestions) || selectedSuggestions.length === 0) {
      throw new Error("No selected suggestions provided");
    }

    // Re-validate defensively in case client state is stale
    const valid = validateSuggestions(selectedSuggestions);
    if (!valid || valid.length === 0) {
      throw new Error("No valid suggestions to apply after validation");
    }

    const result = applySuggestionsToDoc(valid);
    Logger.log("applySelectedSuggestions: applied " + result.applied + " suggestions");
    return { 
      success: true, 
      count: result.applied,
      skipped: result.skipped || 0
    };
  } catch (error) {
    Logger.log("applySelectedSuggestions ERROR: " + error.toString());
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

    const prompt = `Rewrite this resume bullet to be stronger, concise, and achievement-driven while preserving truthfulness. Do NOT add line breaks or new bullets. Keep length <= current length. Use strong action verbs, ATS-friendly phrasing, and STAR cues without inventing new skills/tools. Return ONLY the improved one-line version, no explanations.

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

// Parse suggestions from OpenAI response. Throws error with details if parsing fails.
function parseSuggestions(response) {
  if (!response || typeof response !== "string" || response.trim() === "") {
    throw new Error("Empty or invalid response from OpenAI. Expected JSON array.");
  }

  try {
    // Try to extract JSON array from response (in case there's extra prose)
    const jsonMatch = response.match(/\[[\s\S]*\]/);
    if (jsonMatch) {
      const parsed = JSON.parse(jsonMatch[0]);
      if (!Array.isArray(parsed)) {
        throw new Error("Parsed JSON is not an array. Expected array of suggestion objects.");
      }
      return parsed;
    }
    
    // Try parsing entire response as JSON
    const parsed = JSON.parse(response);
    if (!Array.isArray(parsed)) {
      throw new Error("Parsed JSON is not an array. Expected array of suggestion objects.");
    }
    return parsed;
  } catch (error) {
    Logger.log("Error parsing suggestions: " + error.toString());
    Logger.log("Response content (first 500 chars): " + response.substring(0, 500));
    
    // Fail loudly with detailed error message
    const errorMsg = error.message || error.toString();
    throw new Error(`Failed to parse suggestions as JSON: ${errorMsg}. Response may not be valid JSON or may be missing the expected array format.`);
  }
}

// Apply suggestions to document. Operates on text elements only (not paragraphs) to preserve structure.
// Handles single text-element matches; skips if text spans multiple elements (edge case).
function applySuggestionsToDoc(suggestions) {
  if (!suggestions || suggestions.length === 0) {
    throw new Error("No suggestions to apply");
  }
  
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  let appliedCount = 0;
  const skippedReasons = [];
  
  suggestions.forEach((suggestion, index) => {
    try {
      const original = suggestion.original;
      const replacement = suggestion.suggestion;
      
      if (!original || !replacement) {
        skippedReasons.push(`Suggestion ${index + 1}: Missing original or suggestion text`);
        return;
      }
      
      // Try exact match first (most common case)
      let found = body.findText(original);
      
      // If not found, try with normalized whitespace (handles minor formatting differences)
      if (!found) {
        const normalized = original.replace(/\s+/g, " ").trim();
        found = body.findText(normalized);
      }
      
      if (!found) {
        const reason = `Suggestion ${index + 1}: Text not found in document. Original: "${original.substring(0, 50)}..."`;
        Logger.log("Skipping - " + reason);
        skippedReasons.push(reason);
        return;
      }
      
      const element = found.getElement();
      
      // Only operate on TEXT elements to avoid touching paragraph structure
      if (element.getType() !== DocumentApp.ElementType.TEXT) {
        const reason = `Suggestion ${index + 1}: Match found in non-text element (${element.getType()})`;
        Logger.log("Skipping - " + reason);
        skippedReasons.push(reason);
        return;
      }

      const textElement = element.asText();
      const start = found.getStartOffset();
      const end = found.getEndOffsetInclusive();
      
      // Verify the matched text actually matches (findText can be fuzzy with special chars)
      const matchedText = textElement.getText().substring(start, end + 1);
      if (matchedText !== original && matchedText.replace(/\s+/g, " ").trim() !== original.replace(/\s+/g, " ").trim()) {
        const reason = `Suggestion ${index + 1}: Matched text differs from original`;
        Logger.log("Skipping - " + reason);
        skippedReasons.push(reason);
        return;
      }

      // Capture formatting from the ORIGINAL text range (not just start position)
      // Sample multiple positions to detect what formatting the original text actually had
      const originalLength = end - start + 1;
      const samplePositions = [
        start,
        Math.floor(start + originalLength / 2),
        end
      ].filter(pos => pos >= start && pos <= end);
      
      // Get attributes from each sample position
      const sampleAttributes = samplePositions.map(pos => textElement.getAttributes(pos));
      
      // Determine which attributes were consistently present in the original text
      // Only preserve formatting that was actually in the original (not inherited from surrounding text)
      const preservedAttributes = {};
      if (sampleAttributes.length > 0) {
        const firstAttrs = sampleAttributes[0];
        Object.keys(firstAttrs).forEach(key => {
          // For style attributes (ITALIC, BOLD), only preserve if ALL samples had it
          // This prevents inheriting formatting from adjacent text
          if (key === DocumentApp.Attribute.ITALIC || key === DocumentApp.Attribute.BOLD) {
            const allHaveIt = sampleAttributes.every(attrs => attrs[key] === true);
            if (allHaveIt) {
              preservedAttributes[key] = true;
            } else {
              // Explicitly set to false to clear inherited formatting
              preservedAttributes[key] = false;
            }
          } else {
            // For other attributes (font, size, etc.), use the first sample's value
            if (firstAttrs[key] !== null) {
              preservedAttributes[key] = firstAttrs[key];
            }
          }
        });
      }

      // Replace text content only; never insert paragraph breaks or modify paragraph structure
      textElement.deleteText(start, end);
      textElement.insertText(start, replacement);

      // Apply only the formatting that was in the original text (prevents inheriting italics from above)
      const newEnd = start + replacement.length - 1;
      Object.keys(preservedAttributes).forEach(key => {
        textElement.setAttributes(start, newEnd, { [key]: preservedAttributes[key] });
      });

      appliedCount++;
      Logger.log(`Applied suggestion ${index + 1}: "${original.substring(0, 30)}..." -> "${replacement.substring(0, 30)}..."`);
    } catch (error) {
      const reason = `Suggestion ${index + 1}: Error during application - ${error.toString()}`;
      Logger.log("Error applying suggestion: " + reason);
      skippedReasons.push(reason);
    }
  });
  
  // Log summary
  if (skippedReasons.length > 0) {
    Logger.log(`Applied ${appliedCount}/${suggestions.length} suggestions. Skipped: ${skippedReasons.join("; ")}`);
  }
  
  if (appliedCount === 0) {
    const errorMsg = skippedReasons.length > 0 
      ? `Could not apply any suggestions. Reasons: ${skippedReasons.join("; ")}`
      : "Could not find any matching text to replace. Make sure the original text matches exactly.";
    throw new Error(errorMsg);
  }
  
  // Return info about what was skipped (for future UI enhancement)
  return { applied: appliedCount, skipped: skippedReasons.length };
}

// Analyze resume match to job description: returns match % and missing skills/tools.
// Uses OpenAI to compare resume content with JD requirements.
function analyzeResumeMatch(resumeText, jdText) {
  Logger.log("analyzeResumeMatch: Starting analysis");
  
  const apiKey = getApiKey();
  if (!apiKey) {
    throw new Error("OpenAI API key not found");
  }

  const prompt = `You are a resume analysis assistant. Analyze how well the resume matches the job description.

IMPORTANT RULES:
- Do NOT invent or suggest skills/tools that aren't explicitly mentioned in the job description.
- Only list skills/tools that are CLEARLY stated in the job description but NOT found in the resume.
- Be conservative: if unsure whether something is in the resume, assume it is (don't list it as missing).
- Return ONLY valid JSON, no prose.

Analyze:
1. Overall match percentage (0-100): How well does the resume align with the job requirements?
2. Missing skills/tools: List specific skills, technologies, or tools mentioned in the JD that are NOT present in the resume. Only include items that are explicitly mentioned in the JD.

Return JSON in this exact format:
{
  "match_percentage": <number 0-100>,
  "missing_items": {
    "skills": ["skill1", "skill2"],
    "tools": ["tool1", "tool2"],
    "technologies": ["tech1", "tech2"]
  },
  "notes": "Brief explanation of match percentage (1-2 sentences)"
}

RESUME:
${resumeText}

JOB DESCRIPTION:
${jdText}`;

  try {
    const response = callOpenAI(prompt);
    Logger.log("analyzeResumeMatch: Got response, parsing");
    
    // Extract JSON from response
    const jsonMatch = response.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      throw new Error("No JSON found in response");
    }
    
    const analysis = JSON.parse(jsonMatch[0]);
    
    // Validate structure
    if (typeof analysis.match_percentage !== "number" || analysis.match_percentage < 0 || analysis.match_percentage > 100) {
      throw new Error("Invalid match_percentage in response");
    }
    
    if (!analysis.missing_items || typeof analysis.missing_items !== "object") {
      analysis.missing_items = { skills: [], tools: [], technologies: [] };
    }
    
    Logger.log("analyzeResumeMatch: Success - " + analysis.match_percentage + "% match");
    return analysis;
  } catch (error) {
    Logger.log("analyzeResumeMatch ERROR: " + error.toString());
    throw new Error("Failed to analyze resume match: " + error.message);
  }
}

// Validate suggestions for length and line-safety before applying.
function validateSuggestions(suggestions) {
  if (!suggestions || suggestions.length === 0) return [];
  
  const valid = [];
  
  suggestions.forEach((s, index) => {
    if (!s || !s.original || !s.suggestion) {
      Logger.log("Rejecting suggestion " + index + " (missing original/suggestion)");
      return;
    }

    // Reject if suggestion contains line breaks
    if (/\r|\n/.test(s.suggestion)) {
      Logger.log("Rejecting suggestion " + index + " (contains line break)");
      return;
    }

    // Compute/fallback counts if not provided
    const originalCount = typeof s.original_char_count === "number" ? s.original_char_count : s.original.length;
    const suggestedCount = typeof s.suggested_char_count === "number" ? s.suggested_char_count : s.suggestion.length;

    // Enforce character rules
    if (originalCount <= 123 && suggestedCount > originalCount) {
      Logger.log("Rejecting suggestion " + index + " (suggestion longer than original constraint)");
      return;
    }
    if (originalCount > 123 && suggestedCount > Math.floor(originalCount * 1.15)) {
      Logger.log("Rejecting suggestion " + index + " (exceeds 115% of original)");
      return;
    }

    valid.push({
      ...s,
      original_char_count: originalCount,
      suggested_char_count: suggestedCount
    });
  });

  return valid;
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