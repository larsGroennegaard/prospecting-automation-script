
/**
 * Ensures the Content Library sheet has the correct header.
 */
function ensureContentLibraryHeader_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONTENT_LIB_SHEET);
  if (!sh) throw new Error(`Missing sheet: ${CONTENT_LIB_SHEET}`);
  const expected = ["URL", "Title", "Description / Key Takeaway", "Target Persona"];
  if (sh.getLastRow() > 0) sh.clear(); // Clear the sheet before adding headers
  sh.appendRow(expected);
  SpreadsheetApp.flush();
  return sh;
}

/**
 * Main function to orchestrate the blog cataloging process.
 */
function cl_buildLibrary() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Rebuild Content Library?',
    'This will clear the "Content Library" sheet and rebuild it by scraping and analyzing your blog posts. This can take several minutes. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const sheet = ensureContentLibraryHeader_();

  // You can add more blog index pages here if needed
  const blogIndexPages = [
    "https://dreamdata.io/blog",
    "https://dreamdata.io/blog?offset=1744801352826"
  ];

  try {
    const allPostUrls = cl_getAllPostUrls_(blogIndexPages);
    Logger.log(`Found ${allPostUrls.length} unique blog post URLs.`);

    if (allPostUrls.length === 0) {
      throw new Error("Could not find any blog post URLs. The website's design might have changed.");
    }

    for (const postUrl of allPostUrls) {
      Logger.log(`Processing: ${postUrl}`);
      const html = UrlFetchApp.fetch(postUrl, { muteHttpExceptions: true }).getContentText();
      const contentData = cl_getPostContentAndTitle_(html);
      
      if (contentData && contentData.content.length > 50) {
        const analysis = cl_analyzeContentWithGemini_(contentData.title, contentData.content);
        if (analysis) {
          const [analyzedTitle, description, persona] = analysis;
          sheet.appendRow([postUrl, contentData.title, description, persona]);
        }
      } else {
         Logger.log(`Skipping ${postUrl} due to insufficient content.`);
      }
      Utilities.sleep(1000); // Pause to be respectful to the server
    }

    ui.alert("Success!", `The Content Library has been created with ${allPostUrls.length} posts.`, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert("An Error Occurred", e.message, ui.ButtonSet.OK);
    Logger.log(e);
  }
}

/**
 * Scrapes blog index pages to find all individual post URLs.
 * @param {string[]} urls - An array of blog index page URLs.
 * @returns {string[]} A unique array of blog post URLs.
 */
function cl_getAllPostUrls_(urls) {
  const postUrls = new Set();
  const regex = /<a href="(\/blog\/[^"]+)"/g;

  urls.forEach(url => {
    try {
      const html = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText();
      let match;
      while ((match = regex.exec(html)) !== null) {
        if (!match[1].includes('/category/')) {
          postUrls.add("https://dreamdata.io" + match[1]);
        }
      }
    } catch (e) {
      Logger.log(`Failed to fetch or parse URL: ${url}. Error: ${e.message}`);
    }
  });
  return Array.from(postUrls);
}

/**
 * Extracts the title and clean text content from a blog post's HTML.
 * @param {string} html - The HTML content of a blog post.
 * @returns {Object|null} An object with title and content, or null.
 */
function cl_getPostContentAndTitle_(html) {
  try {
    const titleMatch = html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/);
    const title = titleMatch ? titleMatch[1].trim().replace(/&nbsp;/g, ' ') : "Title not found";

    const contentMatch = html.match(/<div class="blog-item-content-wrapper"[\s\S]*?>([\s\S]*?)<\/section>/);
    let content = "Content not found";
    if (contentMatch) {
      content = contentMatch[1].replace(/<[^>]*>/g, ' ').replace(/\s\s+/g, ' ').trim();
    }
    
    return { title, content };
  } catch (e) {
    Logger.log(`Failed to get content. Error: ${e.message}`);
    return null;
  }
}

/**
 * Sends content to the Gemini API for analysis and returns a parsed CSV row.
 * @param {string} title - The title of the blog post.
 * @param {string} content - The text content of the blog post.
 * @returns {string[]|null} An array of [title, description, persona].
 */
function cl_analyzeContentWithGemini_(title, content) {
  const apiKey = cfg_('GEMINI_API_KEY'); // Using the cfg_ utility
  const url = `https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const truncatedContent = content.substring(0, 15000);

  const prompt = `
    You are an expert B2B Content Analyst for Dreamdata, a B2B GTM Attribution Platform.
    TARGET PERSONAS: CMO / VP Marketing, VP Demand Generation, Head of Marketing Ops, Head of Performance Marketing.

    ANALYZE THE FOLLOWING BLOG POST:
    Title: "${title}"
    Content: "${truncatedContent}"

    YOUR TASK:
    1.  Description / Key Takeaway: Write a single, crisp sentence describing the key problem the article solves.
    2.  Target Persona: Identify the primary target persona from the list above.

    OUTPUT FORMAT:
    You MUST respond with a single line of CSV with THREE fields, each enclosed in double quotes: "Title","Description / Key Takeaway","Target Persona".
    The title in your output MUST MATCH the input title exactly.
  `;

  const payload = { 
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": { "maxOutputTokens": 512 }
  };
  const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(payload), 'muteHttpExceptions': true };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    const jsonResponse = JSON.parse(responseText);
    
    if (jsonResponse.candidates && jsonResponse.candidates[0].content.parts[0].text) {
      let result = jsonResponse.candidates[0].content.parts[0].text.trim();
      // Simple CSV parser for "field1","field2","field3"
      return result.split('","').map(s => s.replace(/"/g, ''));
    } else {
      Logger.log(`Invalid API response for title "${title}": ${responseText}`);
      return [title, "Analysis Failed: Invalid API response", "N/A"];
    }
  } catch (e) {
    Logger.log(`API call failed for title "${title}". Error: ${e.message}`);
    return [title, `Analysis Failed: ${e.message}`, "N/A"];
  }
}