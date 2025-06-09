import fetch from 'node-fetch';
import { getAioLogger, strToArray } from '../utils.js';
import AppConfig from '../appConfig.js';
import HelixUtils from '../helixUtils.js';

async function main(params) {
    const logger = getAioLogger('find-fragments', params.LOG_LEVEL || 'info');
    // Convert sourcePaths to array if it's a string
    const sourcePaths = strToArray(params.sourcePaths);
    if (!Array.isArray(sourcePaths) || sourcePaths.length === 0) {
        return {
            statusCode: 400,
            body: {
                error: 'sourcePaths must be a non-empty array or comma-separated string'
            }
        };
    }

    const appConfig = new AppConfig(params);
    const helixUtils = new HelixUtils(appConfig);
    const fragmentLinks = new Set();
    const processedPaths = new Set(); // Tracking processed paths to avoid infinite loops

    // Process all AEM URLs in parallel
    const aemPaths = sourcePaths.filter((path) => path.includes('aem.page'));

    const processPath = async (originalPath, isFragment = false) => {
        // Create a copy of the path to avoid modifying the parameter
        let pathToProcess = originalPath;

        // Skip if already processed to avoid infinite loops
        if (processedPaths.has(pathToProcess)) {
            return [];
        }
        processedPaths.add(pathToProcess);

        // Fetch the markdown content
        const options = {};
        // Passing isGraybox param true to fetch graybox Hlx Admin API Key
        const grayboxHlxAdminApiKey = helixUtils.getAdminApiKey(false);
        if (grayboxHlxAdminApiKey) {
            options.headers = new fetch.Headers();
            options.headers.append('Authorization', `token ${grayboxHlxAdminApiKey}`);
        }

        // Add .md extension if not already present
        if (!pathToProcess.endsWith('.md')) {
            pathToProcess += '.md';
        }

        const response = await fetch(`${pathToProcess}`, options);
        const content = await response.text();
        logger.info(`Content from ${isFragment ? 'fragment' : 'sharepoint'} in find-fragments: ${content.substring(0, 500)}...`);

        // Find fragment links in content using angle bracket format
        // Pattern matches: <https://...aem.page/.../fragments/...>
        const fragmentMatches = content.match(/<https:\/\/[^>]*aem\.page[^>]*\/fragments\/[^>]*>/g) || [];
        const pathFragmentLinks = [];

        fragmentMatches.forEach((match) => {
            // Remove angle brackets to get the clean URL
            const cleanUrl = match.slice(1, -1);
            pathFragmentLinks.push(cleanUrl);
        });

        logger.info(`Found ${fragmentMatches.length} fragment links in ${originalPath}`);
        // Recursively process each fragment found using Promise.all
        const recursiveFragmentPromises = pathFragmentLinks.map(async (fragmentUrl) => {
            try {
                return await processPath(fragmentUrl, true);
            } catch (error) {
                logger.error(`Error processing fragment ${fragmentUrl}: ${error.message}`);
                return [];
            }
        });

        const recursiveResults = await Promise.all(recursiveFragmentPromises);
        const flattenedRecursiveResults = recursiveResults.flat();

        // Return both current level fragments and nested fragments
        return [...pathFragmentLinks, ...flattenedRecursiveResults];
    };

    // Process all AEM paths in parallel
    const results = await Promise.all(aemPaths.map((path) => processPath(path)));

    // Add all found fragment links to the set
    results.forEach((pathLinks) => {
        pathLinks.forEach((link) => fragmentLinks.add(link));
    });

    logger.info(`Found fragment links: ${Array.from(fragmentLinks).join(', ')}`);

    return {
        statusCode: 200,
        body: {
            fragmentLinks: Array.from(fragmentLinks)
        }
    };
}

export { main };
