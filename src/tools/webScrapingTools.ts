// Web Scraping Tool - Hybrid Mode
export interface WebScrapingOptions {
  waitForSelector?: string;
  timeout?: number;
  extractImages?: boolean;
  extractLinks?: boolean;
  textOnly?: boolean;
}

interface WebScrapingResult {
  success: boolean;
  data?: any;
  error?: string;
  instructions?: string;
  mcpCommands?: string[];
}

/**
 * 🕷️ Web Content Scraping Tool - Hybrid Mode: Uses playwright-mcp to scrape webpage content
 */
export async function scrapeWebContent(
  url: string,
  options: WebScrapingOptions = {}
): Promise<WebScrapingResult> {
  try {
    // Validate URL format
    try {
      new URL(url);
    } catch (error) {
      return {
        success: false,
        error: `Invalid URL format: ${url}`,
      };
    }

    // Build playwright-mcp commands
    const playwrightCommands = [
      `browser_launch({ headless: true })`,
      `browser_navigate("${url}")`,
      `browser_wait_for({ time: 10 })`,
    ];

    if (options.waitForSelector) {
      playwrightCommands.push(`browser_wait_for({ selector: "${options.waitForSelector}" })`);
    }

    const extractionCommand = options.textOnly ? `browser_get_text()` : `browser_get_html()`;

    playwrightCommands.push(extractionCommand);

    const instructions = [
      '🔧 Web Content Scraping Workflow:',
      '',
      '1. Run the following playwright-mcp commands:',
      ...playwrightCommands.map((cmd, i) => `   ${i + 1}. ${cmd}`),
      '',
      '2. Continue processing after retrieving the page content',
    ].join('\n');

    return {
      success: true,
      instructions,
      mcpCommands: playwrightCommands,
    };
  } catch (error) {
    return {
      success: false,
      error: error instanceof Error ? error.message : 'Unknown error',
    };
  }
}

/**
 * 📊 Structured Data Scraping Tool
 */
export async function scrapeStructuredData(
  url: string,
  selector: string,
  options: WebScrapingOptions = {}
): Promise<WebScrapingResult> {
  try {
    // Validate URL format
    try {
      new URL(url);
    } catch (error) {
      return {
        success: false,
        error: `Invalid URL format: ${url}`,
      };
    }

    const instructions = [
      '📊 Structured Data Scraping Workflow:',
      '',
      `1. Navigate to ${url}`,
      `2. Wait for element ${selector}`,
      `3. Extract structured data`,
    ].join('\n');

    return {
      success: true,
      instructions,
      mcpCommands: [
        `browser_launch({ headless: true })`,
        `browser_navigate("${url}")`,
        `browser_wait_for({ selector: "${selector}" })`,
        `browser_get_text("${selector}")`,
      ],
    };
  } catch (error) {
    return {
      success: false,
      error: error instanceof Error ? error.message : 'Unknown error',
    };
  }
}

// MCP Tool Definitions
export const WEB_SCRAPING_TOOL = {
  name: 'scrape_web_content',
  description: '🕷️ Web content scraping - Use Playwright Chromium to scrape webpage content.',
  inputSchema: {
    type: 'object',
    properties: {
      url: {
        type: 'string',
        description: 'Webpage URL to scrape',
      },
      options: {
        type: 'object',
        properties: {
          waitForSelector: { type: 'string', description: 'CSS selector to wait for' },
          timeout: { type: 'number', description: 'Timeout in milliseconds' },
          textOnly: { type: 'boolean', description: 'Extract only plain text' },
        },
      },
    },
    required: ['url'],
  },
};

export const STRUCTURED_DATA_TOOL = {
  name: 'scrape_structured_data',
  description:
    '📊 Structured data scraping - Scrape structured data from webpages using a CSS selector.',
  inputSchema: {
    type: 'object',
    properties: {
      url: {
        type: 'string',
        description: 'Webpage URL to scrape',
      },
      selector: {
        type: 'string',
        description: 'CSS selector',
      },
      options: {
        type: 'object',
        properties: {
          timeout: { type: 'number', description: 'Timeout in milliseconds' },
        },
      },
    },
    required: ['url', 'selector'],
  },
};
