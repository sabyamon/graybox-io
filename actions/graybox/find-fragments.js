import fetch from 'node-fetch';
import util from 'util';
import { getAioLogger, strToArray } from '../utils.js';
import Sharepoint from '../sharepoint.js';
import AppConfig from '../appConfig.js';
import HelixUtils from '../helixUtils.js';

async function main(params) {
    // create a Logger
    const logger = getAioLogger('find-fragments', params.LOG_LEVEL || 'info');
    logger.info(`Params in Find Fragments: ${JSON.stringify(params)}`);
    logger.info('Starting find fragments operation');

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
    logger.info(`AppConfig in find-fragments: ${JSON.stringify(appConfig)}`);
    const helixUtils = new HelixUtils(appConfig);
    const sharepoint = new Sharepoint(appConfig);
    logger.info(`Sharepoint in find-fragments: ${JSON.stringify(sharepoint)}`);
    const fragmentLinks = new Set();

    // Process all AEM URLs in parallel
    const aemPaths = sourcePaths.filter((path) => path.includes('aem.page'));

    const processPath = async (path) => {
        // Fetch the markdown content
        const options = {};
        // Passing isGraybox param true to fetch graybox Hlx Admin API Key
        const grayboxHlxAdminApiKey = helixUtils.getAdminApiKey(false);
        logger.info(`Graybox Hlx Admin API Key in find-fragments: ${grayboxHlxAdminApiKey}`);
        if (grayboxHlxAdminApiKey) {
            options.headers = new fetch.Headers();
            options.headers.append('Authorization', `token ${grayboxHlxAdminApiKey}`);
        }
        path += '.md';
        logger.info(`Fetching content for in find-fragments: ${path}`);
        logger.info(`Options in find-fragments: ${JSON.stringify(options)}`);
        const response = await sharepoint.fetchWithRetry(`${path}`, options);
        const fileDataResponse = await sharepoint.getFileData('/sabya/drafts/sabya-doc-1', false);
        logger.info(`File Data Response in find-fragments: ${JSON.stringify(fileDataResponse)}`);
        logger.info(`Response from sharepoint in find-fragments: ${util.inspect(response, { depth: null, colors: true })}`);
        const content = await response.text();
        logger.info(`Content from sharepoint in find-fragments: ${content}`);

        // Find fragment links in content
        const fragmentMatches = content.match(/\[.*?\]\(.*?\/fragments\/.*?\)/g) || [];
        const pathFragmentLinks = [];
        fragmentMatches.forEach((match) => {
            const linkMatch = match.match(/\((.*?)\)/);
            if (linkMatch && linkMatch[1]) {
                pathFragmentLinks.push(linkMatch[1]);
            }
        });

        logger.info(`Found ${fragmentMatches.length} fragment links in ${path}`);
        return pathFragmentLinks;
    };

    // Process all AEM paths in parallel
    const results = await Promise.all(aemPaths.map(processPath));

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
