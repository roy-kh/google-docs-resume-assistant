// ============================================
// CONFIGURATION: Set your API keys here
// ============================================
// Replace 'YOUR_API_KEY_HERE' with your actual API keys before use.
// Do NOT commit real keys to source control.
const OPENAI_API_KEY = "YOUR_API_KEY_HERE";
const PERPLEXITY_API_KEY = "YOUR_PERPLEXITY_API_KEY_HERE";

function onOpen() {
  DocumentApp.getUi()
    .createMenu("Cover Letter Assistant")
    .addItem("Open Assistant", "showSidebar")
    .addItem("Authorize Script", "authorizeScript")
    .addToUi();
}

// This function triggers authorization - MUST be run from Apps Script editor first time
function authorizeScript() {
  try {
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

function testAuthorization() {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty("OPENAI_KEY");
  Logger.log("API Key found: " + (key ? "Yes (starts with " + key.substring(0, 7) + "...)" : "No"));
  return "Authorization successful! API key " + (key ? "found" : "not found");
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("sidebar")
    .setTitle("Cover Letter Assistant");
  DocumentApp.getUi().showSidebar(html);
}

// Helper function to get API key
function getApiKey() {
  if (OPENAI_API_KEY && OPENAI_API_KEY !== "YOUR_API_KEY_HERE") return OPENAI_API_KEY;
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

// Helper function to get Perplexity API key
function getPerplexityApiKey() {
  if (PERPLEXITY_API_KEY && PERPLEXITY_API_KEY !== "YOUR_PERPLEXITY_API_KEY_HERE") return PERPLEXITY_API_KEY;
  try {
    const docKey = PropertiesService.getDocumentProperties().getProperty("PERPLEXITY_KEY");
    if (docKey) return docKey;
    const scriptKey = PropertiesService.getScriptProperties().getProperty("PERPLEXITY_KEY");
    if (scriptKey) return scriptKey;
  } catch (e) {
    Logger.log("PropertiesService not accessible for Perplexity: " + e.toString());
  }
  return null;
}

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

// Extract resume highlights relevant to the job description
function extractResumeHighlights(resumeText, jdText) {
  Logger.log("extractResumeHighlights: Starting");
  
  const apiKey = getApiKey();
  if (!apiKey) {
    throw new Error("OpenAI API key not found");
  }

  const prompt = `Extract 3-5 most relevant resume bullets or experiences that align with this job description. Return a concise summary (200-300 words max) focusing on achievements, skills, and experiences that match JD requirements. Format as a brief narrative or bullet points.

RESUME:
${resumeText}

JOB DESCRIPTION:
${jdText}`;

  try {
    const response = callOpenAI(prompt);
    Logger.log("extractResumeHighlights: Got response");
    return response.trim();
  } catch (error) {
    Logger.log("extractResumeHighlights ERROR: " + error.toString());
    throw new Error("Failed to extract resume highlights: " + error.message);
  }
}

// Get company overview from Perplexity API
function getCompanyOverview(companyName) {
  Logger.log("getCompanyOverview: Starting for " + companyName);
  
  const apiKey = getPerplexityApiKey();
  if (!apiKey) {
    Logger.log("Perplexity API key not found, skipping company overview");
    return null;
  }

  if (!companyName || companyName.trim() === "") {
    return null;
  }

  const url = "https://api.perplexity.ai/chat/completions";
  
  const payload = {
    model: "llama-3.1-sonar-large-128k-online",
    messages: [
      {
        role: "system",
        content: "You are a helpful assistant that provides concise company overviews for job applications. Focus on company culture, values, recent news, and what makes them unique. Keep response to 200-300 words."
      },
      {
        role: "user",
        content: `Provide a brief overview of ${companyName} focusing on company culture, values, recent developments, and what makes them an attractive employer.`
      }
    ],
    temperature: 0.2,
    max_tokens: 500
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + apiKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      Logger.log("Perplexity API error: " + responseText);
      return null; // Non-fatal, continue without company overview
    }

    const data = JSON.parse(responseText);
    
    if (!data.choices || !data.choices[0] || !data.choices[0].message) {
      Logger.log("Unexpected Perplexity response format");
      return null;
    }

    const overview = data.choices[0].message.content;
    Logger.log("getCompanyOverview: Success");
    return overview.trim();
  } catch (error) {
    Logger.log("getCompanyOverview ERROR: " + error.toString());
    return null; // Non-fatal, continue without company overview
  }
}

// Main function to process cover letter
function processCoverLetter(jdText, companyName, resumeText) {
  Logger.log("=== processCoverLetter START ===");
  
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      throw new Error("OpenAI API key not found. Please set OPENAI_API_KEY constant at the top of Code.gs");
    }
    
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const coverLetterText = body.getText();

    if (!jdText || jdText.trim() === "") {
      throw new Error("Please provide a job description");
    }

    // Step 1: Get company overview (non-fatal if fails)
    let companyOverview = null;
    if (companyName && companyName.trim()) {
      try {
        companyOverview = getCompanyOverview(companyName.trim());
      } catch (error) {
        Logger.log("Company overview fetch failed (non-fatal): " + error.toString());
      }
    }

    // Step 2: Extract resume highlights if resume provided
    let resumeHighlights = "";
    if (resumeText && resumeText.trim()) {
      try {
        resumeHighlights = extractResumeHighlights(resumeText.trim(), jdText);
        Logger.log("Resume highlights extracted: " + resumeHighlights.length + " chars");
      } catch (error) {
        Logger.log("Resume highlights extraction failed (non-fatal): " + error.toString());
        // Continue without resume highlights
      }
    }

    // Step 3: Build prompt for cover letter suggestions
    let contextParts = [];
    if (resumeHighlights) {
      contextParts.push(`RELEVANT RESUME HIGHLIGHTS (for reference):\n${resumeHighlights}`);
    }
    if (companyOverview) {
      contextParts.push(`COMPANY OVERVIEW:\n${companyOverview}`);
    }

    const prompt = `You are a cover letter editing assistant. Return ONLY valid JSON (no prose). Optimize for clarity, professionalism, and alignment with the job description while preserving the candidate's authentic voice.

KEYWORD OPTIMIZATION: When improving paragraphs, actively incorporate relevant keywords, technologies, and terminology from the job description where they naturally fit and truthfully represent the candidate's experience.

CRITICAL: Only suggest changes for paragraphs that NEED improvement. If a paragraph is already strong, clear, and well-written, DO NOT include it in your suggestions. Only suggest changes when there's meaningful improvement to be made (weak phrasing, missing keywords, poor alignment with JD, unclear value proposition, etc.).

HARD RULES:
- Do not add new paragraphs or remove existing paragraphs.
- Keep formatting consistent; preserve tense/person/voice.
- Suggestions should maintain similar length to original (within 20%).
- Suggestion text must preserve paragraph structure (can contain line breaks for paragraphs).
- Self-report character counts for original and suggestion (count all characters).
- Return an EMPTY array [] if no paragraphs need improvement.

EXPECTED JSON ARRAY FORMAT ONLY:
[
  {
    "original": "exact text from cover letter paragraph",
    "suggestion": "improved version (can contain line breaks for paragraph structure)",
    "reason": "why this helps (clarity/alignment/keyword optimization, truthful)",
    "original_char_count": <number>,
    "suggested_char_count": <number>
  }
]

COVER LETTER:
${coverLetterText}

JOB DESCRIPTION:
${jdText}

${contextParts.length > 0 ? contextParts.join("\n\n") : ""}`;

    Logger.log("Step 4: Calling OpenAI");
    const response = callOpenAI(prompt);
    Logger.log("Step 5: Parsing suggestions");
    
    let suggestions;
    try {
      suggestions = parseSuggestions(response);
    } catch (parseError) {
      Logger.log("Parse error details: " + parseError.toString());
      return {
        success: false,
        error: "parse",
        message: parseError.message || parseError.toString(),
        count: 0,
        suggestions: []
      };
    }
    
    Logger.log("Step 6: Validating suggestions");
    const validSuggestions = validateCoverLetterSuggestions(suggestions);
    
    Logger.log("Step 7: Returning suggestions to client");
    return { 
      success: true, 
      count: validSuggestions.length, 
      suggestions: validSuggestions,
      companyOverview: companyOverview
    };
  } catch (error) {
    Logger.log("ERROR in processCoverLetter: " + error.toString());
    Logger.log("Error stack: " + (error.stack || "No stack trace"));
    throw error;
  }
}

// Apply selected cover letter suggestions
function applySelectedCoverLetterSuggestions(selectedSuggestions) {
  Logger.log("applySelectedCoverLetterSuggestions: start");
  try {
    if (!selectedSuggestions || !Array.isArray(selectedSuggestions) || selectedSuggestions.length === 0) {
      throw new Error("No selected suggestions provided");
    }

    const valid = validateCoverLetterSuggestions(selectedSuggestions);
    if (!valid || valid.length === 0) {
      throw new Error("No valid suggestions to apply after validation");
    }

    const result = applyCoverLetterSuggestionsToDoc(valid);
    Logger.log("applySelectedCoverLetterSuggestions: applied " + result.applied + " suggestions");
    return { 
      success: true, 
      count: result.applied,
      skipped: result.skipped || 0
    };
  } catch (error) {
    Logger.log("applySelectedCoverLetterSuggestions ERROR: " + error.toString());
    throw error;
  }
}

// Validate cover letter suggestions (more lenient than resume - allows paragraphs)
function validateCoverLetterSuggestions(suggestions) {
  if (!suggestions || suggestions.length === 0) return [];
  
  const valid = [];
  
  suggestions.forEach((s, index) => {
    if (!s || !s.original || !s.suggestion) {
      Logger.log("Rejecting suggestion " + index + " (missing original/suggestion)");
      return;
    }

    // Compute/fallback counts if not provided
    const originalCount = typeof s.original_char_count === "number" ? s.original_char_count : s.original.length;
    const suggestedCount = typeof s.suggested_char_count === "number" ? s.suggested_char_count : s.suggestion.length;

    // Enforce character rules (20% variance allowed for paragraphs)
    const maxAllowed = Math.floor(originalCount * 1.2);
    const minAllowed = Math.floor(originalCount * 0.8);
    
    if (suggestedCount > maxAllowed) {
      Logger.log("Rejecting suggestion " + index + " (exceeds 120% of original)");
      return;
    }
    if (suggestedCount < minAllowed) {
      Logger.log("Rejecting suggestion " + index + " (below 80% of original)");
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

// Apply cover letter suggestions to document (paragraph-level)
function applyCoverLetterSuggestionsToDoc(suggestions) {
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
      
      // Try exact match first
      let found = body.findText(original);
      
      // If not found, try with normalized whitespace
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
      
      // For cover letters, we work with paragraphs
      if (element.getType() !== DocumentApp.ElementType.PARAGRAPH && element.getType() !== DocumentApp.ElementType.TEXT) {
        const reason = `Suggestion ${index + 1}: Match found in unsupported element type`;
        Logger.log("Skipping - " + reason);
        skippedReasons.push(reason);
        return;
      }

      // Get the paragraph containing the text
      let paragraph;
      if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
        paragraph = element.asParagraph();
      } else {
        paragraph = element.getParent().asParagraph();
      }

      const start = found.getStartOffset();
      const end = found.getEndOffsetInclusive();
      
      // Get the text element
      const textElement = paragraph.editAsText();
      const matchedText = textElement.getText().substring(start, end + 1);
      
      // Verify match
      if (matchedText !== original && matchedText.replace(/\s+/g, " ").trim() !== original.replace(/\s+/g, " ").trim()) {
        const reason = `Suggestion ${index + 1}: Matched text differs from original`;
        Logger.log("Skipping - " + reason);
        skippedReasons.push(reason);
        return;
      }

      // Capture formatting
      const originalLength = end - start + 1;
      const samplePositions = [
        start,
        Math.floor(start + originalLength / 2),
        end
      ].filter(pos => pos >= start && pos <= end);
      
      const sampleAttributes = samplePositions.map(pos => textElement.getAttributes(pos));
      const preservedAttributes = {};
      
      if (sampleAttributes.length > 0) {
        const firstAttrs = sampleAttributes[0];
        Object.keys(firstAttrs).forEach(key => {
          if (key === DocumentApp.Attribute.ITALIC || key === DocumentApp.Attribute.BOLD) {
            const allHaveIt = sampleAttributes.every(attrs => attrs[key] === true);
            if (allHaveIt) {
              preservedAttributes[key] = true;
            } else {
              preservedAttributes[key] = false;
            }
          } else {
            if (firstAttrs[key] !== null) {
              preservedAttributes[key] = firstAttrs[key];
            }
          }
        });
      }

      // Replace text
      textElement.deleteText(start, end);
      textElement.insertText(start, replacement);

      // Apply formatting
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
  
  if (skippedReasons.length > 0) {
    Logger.log(`Applied ${appliedCount}/${suggestions.length} suggestions. Skipped: ${skippedReasons.join("; ")}`);
  }
  
  if (appliedCount === 0) {
    const errorMsg = skippedReasons.length > 0 
      ? `Could not apply any suggestions. Reasons: ${skippedReasons.join("; ")}`
      : "Could not find any matching text to replace. Make sure the original text matches exactly.";
    throw new Error(errorMsg);
  }
  
  return { applied: appliedCount, skipped: skippedReasons.length };
}

// Parse suggestions from OpenAI response
function parseSuggestions(response) {
  if (!response || typeof response !== "string" || response.trim() === "") {
    throw new Error("Empty or invalid response from OpenAI. Expected JSON array.");
  }

  try {
    const jsonMatch = response.match(/\[[\s\S]*\]/);
    if (jsonMatch) {
      const parsed = JSON.parse(jsonMatch[0]);
      if (!Array.isArray(parsed)) {
        throw new Error("Parsed JSON is not an array. Expected array of suggestion objects.");
      }
      return parsed;
    }
    
    const parsed = JSON.parse(response);
    if (!Array.isArray(parsed)) {
      throw new Error("Parsed JSON is not an array. Expected array of suggestion objects.");
    }
    return parsed;
  } catch (error) {
    Logger.log("Error parsing suggestions: " + error.toString());
    Logger.log("Response content (first 500 chars): " + response.substring(0, 500));
    
    const errorMsg = error.message || error.toString();
    throw new Error(`Failed to parse suggestions as JSON: ${errorMsg}. Response may not be valid JSON or may be missing the expected array format.`);
  }
}

// Call OpenAI API
function callOpenAI(prompt) {
  Logger.log("callOpenAI: Starting");
  
  const API_KEY = getApiKey();
  if (!API_KEY) {
    throw new Error("OpenAI API key not configured. Please use 'Cover Letter Assistant > Configure API Key' to set it up.");
  }
  
  const url = "https://api.openai.com/v1/chat/completions";

  const payload = {
    model: "gpt-4o-mini",
    messages: [
      { role: "system", content: "You are a cover letter editing assistant. Always return valid JSON when requested." },
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

function triggerAuthorizationFromSidebar() {
  try {
    Logger.log("triggerAuthorizationFromSidebar: Starting");
    const key = getApiKey();
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
      message: "Authorization successful! API key found"
    };
  } catch (error) {
    Logger.log("triggerAuthorizationFromSidebar ERROR: " + error.toString());
    throw error;
  }
}

function saveApiKey(key) {
  Logger.log("saveApiKey called");
  try {
    if (!key || key.trim() === "") {
      throw new Error("API key cannot be empty");
    }
    setApiKey(key.trim());
    Logger.log("saveApiKey completed successfully");
    return { success: true, message: "API key saved successfully!" };
  } catch (error) {
    Logger.log("ERROR in saveApiKey: " + error.toString());
    throw error;
  }
}
