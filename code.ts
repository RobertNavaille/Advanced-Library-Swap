// Main Figma Swap Library Plugin (clean template)
console.log('🚀 Plugin starting... Version: No-Legacy-Cleanup-Fix');
import { COMPONENT_KEY_MAPPING as DEFAULT_COMPONENT_KEY_MAPPING, STYLE_KEY_MAPPING as DEFAULT_STYLE_KEY_MAPPING, VARIABLE_ID_MAPPING as DEFAULT_VARIABLE_ID_MAPPING, VARIABLE_KEY_MAPPING as DEFAULT_VARIABLE_KEY_MAPPING, LIBRARY_THUMBNAILS as DEFAULT_LIBRARY_THUMBNAILS } from './keyMapping';
import { copyTextOverrides } from './swapUtils';

// Global mapping variables
let COMPONENT_KEY_MAPPING = DEFAULT_COMPONENT_KEY_MAPPING;
let STYLE_KEY_MAPPING = DEFAULT_STYLE_KEY_MAPPING;
let VARIABLE_ID_MAPPING = DEFAULT_VARIABLE_ID_MAPPING;
let VARIABLE_KEY_MAPPING = DEFAULT_VARIABLE_KEY_MAPPING;
let LIBRARY_THUMBNAILS = DEFAULT_LIBRARY_THUMBNAILS;

// Interface definitions
interface ComponentInfo {
  id: string;
  name: string;  // The variant name (mapping key) - used for swapping
  displayName: string;  // The parent component name - used for UI display
  library: string;
  remote: boolean;
  parentName: string;  // Kept for backward compatibility
  libraryFileId?: string;  // File ID of the library
}

interface TokenInfo {
  id: string;
  name: string;
  type: 'color' | 'typography' | 'spacing' | 'effect';
  value: string;
  library?: string;
}

// Store the scanned frame for later use during swaps
let scannedFrame: FrameNode | null = null;

// Show the UI
figma.showUI(__html__, { width: 480, height: 500, themeColors: true });

// Auto-scan when a frame is selected
figma.on('selectionchange', async () => {
  const selection = figma.currentPage.selection;
  const selectedFrame = selection.find(node => node.type === 'FRAME');
  if (selectedFrame) {
    setTimeout(async () => { await handleScanFrames(); }, 100);
  } else if (selection.length === 0) {
    figma.ui.postMessage({ type: 'SHOW_INITIAL_VIEW' });
  }
});

// Auto-scan on plugin load
(async () => {
  const selection = figma.currentPage.selection;
  const selectedFrame = selection.find(node => node.type === 'FRAME');
  if (selectedFrame) await handleScanFrames();
})();

// Connected Libraries Management
let CONNECTED_LIBRARIES: { 
  name: string; 
  id: string; 
  key: string; 
  type: string;
  lastSynced?: string;
  components?: Record<string, string>;
  styles?: Record<string, string>;
}[] = [];

async function loadConnectedLibraries() {
  console.log('📥 Loading connected libraries...');
  try {
    const stored = await figma.clientStorage.getAsync('connected_libraries');
    console.log('📦 Raw stored libraries:', stored);
    
    if (stored && Array.isArray(stored)) {
      CONNECTED_LIBRARIES = stored;
      console.log(`✅ Loaded ${CONNECTED_LIBRARIES.length} libraries`);
      CONNECTED_LIBRARIES.forEach((lib, index) => {
          const compCount = lib.components ? Object.keys(lib.components).length : 0;
          const styleCount = lib.styles ? Object.keys(lib.styles).length : 0;
          console.log(`   📚 Library [${index}] "${lib.name}": ${compCount} components, ${styleCount} styles.`);
          if (compCount === 0) console.warn(`   ⚠️ Library "${lib.name}" has 0 components! It might not have been fetched correctly.`);
      });
    } else {
      console.log('⚠️ No libraries found in storage or invalid format');
      CONNECTED_LIBRARIES = [];
    }
  } catch (err) {
    console.error('❌ Error loading connected libraries:', err);
    CONNECTED_LIBRARIES = [];
  }

  // Send to UI
  sendConnectedLibraries();
  
  // Update global mappings so scan can find these libraries
  updateMappingsFromConnected();
  
  // Trigger a background refresh to ensure data is up-to-date
  refreshConnectedLibraries();
}

async function refreshConnectedLibraries() {
  console.log('🔄 Refreshing connected libraries in background...');
  const token = await figma.clientStorage.getAsync('figma_access_token');
  if (!token) {
      console.log('⚠️ Cannot refresh libraries: No access token.');
      return;
  }

  let updatedCount = 0;
  for (let i = 0; i < CONNECTED_LIBRARIES.length; i++) {
      const lib = CONNECTED_LIBRARIES[i];
      // Only refresh remote libraries that have a key/id
      if (lib.id && lib.type === 'Remote') {
          try {
              console.log(`   - Refreshing ${lib.name} (${lib.id})...`);
              // Fetch Components
              const compResponse = await fetch(`https://api.figma.com/v1/files/${lib.id}/components`, {
                  headers: { 'X-Figma-Token': token }
              });
              
              let components: Record<string, string> = {};
              if (compResponse.ok) {
                  const compData = await compResponse.json();
                  if (compData.meta && compData.meta.components) {
                      compData.meta.components.forEach((c: any) => {
                          components[c.name] = c.key;
                      });
                  }
              }

              // Fetch Styles
              const styleResponse = await fetch(`https://api.figma.com/v1/files/${lib.id}/styles`, {
                  headers: { 'X-Figma-Token': token }
              });
              
              let styles: Record<string, string> = {};
              if (styleResponse.ok) {
                  const styleData = await styleResponse.json();
                  if (styleData.meta && styleData.meta.styles) {
                      styleData.meta.styles.forEach((s: any) => {
                          styles[s.name] = s.key;
                      });
                  }
              }
              
              // Update the library object
              CONNECTED_LIBRARIES[i] = {
                  ...lib,
                  components: components,
                  styles: styles,
                  lastSynced: new Date().toISOString()
              };
              updatedCount++;
              console.log(`   ✅ Refreshed ${lib.name}: ${Object.keys(components).length} components, ${Object.keys(styles).length} styles`);
              
          } catch (err) {
              console.error(`   ❌ Failed to refresh ${lib.name}:`, err);
          }
      }
  }

  if (updatedCount > 0) {
      saveConnectedLibraries();
      updateMappingsFromConnected();
      sendConnectedLibraries();
      // Trigger a re-scan if we have a selection
      const selection = figma.currentPage.selection;
      if (selection.length > 0) {
          console.log('🔄 Re-scanning after library refresh...');
          await handleScanFrames();
      }
  }
}

function updateMappingsFromConnected() {
  CONNECTED_LIBRARIES.forEach(lib => {
    if (lib.components) {
      COMPONENT_KEY_MAPPING[lib.name] = { ...COMPONENT_KEY_MAPPING[lib.name], ...lib.components };
    }
    if (lib.styles) {
      STYLE_KEY_MAPPING[lib.name] = { ...STYLE_KEY_MAPPING[lib.name], ...lib.styles };
    }
  });
  console.log('🔄 Updated global mappings from connected libraries');
  console.log('Component Mappings Keys:', Object.keys(COMPONENT_KEY_MAPPING));
  Object.keys(COMPONENT_KEY_MAPPING).forEach(lib => {
      console.log(`Library ${lib} has ${Object.keys(COMPONENT_KEY_MAPPING[lib]).length} components`);
  });
}

async function saveConnectedLibraries() {
  try {
    await figma.clientStorage.setAsync('connected_libraries', CONNECTED_LIBRARIES);
    console.log(`💾 Saved ${CONNECTED_LIBRARIES.length} libraries to storage`);
  } catch (err) {
    console.error('❌ Error saving connected libraries:', err);
  }
}

function sendConnectedLibraries() {
  console.log(`📤 Sending ${CONNECTED_LIBRARIES.length} connected libraries to UI`);
  figma.ui.postMessage({
    type: 'LIBRARIES_UPDATED',
    libraries: CONNECTED_LIBRARIES
  });
}

async function handleAddLibraryByLink(link: string) {
  // Extract file key from link
  // Supports:
  // https://www.figma.com/file/ByKey123/Name
  // https://www.figma.com/design/ByKey123/Name
  const match = link.match(/figma\.com\/(?:file|design)\/([a-zA-Z0-9]{22,})/);
  
  if (match && match[1]) {
    const fileKey = match[1];
    console.log('🔗 Extracted file key:', fileKey);
    await handleAddLibraryByKey(fileKey);
  } else {
    figma.notify('Invalid Figma link. Could not extract file key.');
    console.error('Could not extract key from link:', link);
  }
}

async function handleAddLibraryByKey(fileKey: string) {
  console.log(`➕ Adding library by key: ${fileKey}`);
  // 1. Check if we have a token
  const token = await figma.clientStorage.getAsync('figma_access_token');
  
  let name = `Library ${fileKey.substring(0, 6)}`;
  let components: Record<string, string> = {};
  let styles: Record<string, string> = {};
  
  if (token) {
    try {
      // Fetch File Metadata
      const fileResponse = await fetch(`https://api.figma.com/v1/files/${fileKey}`, {
        headers: { 'X-Figma-Token': token }
      });
      
      if (fileResponse.ok) {
        const data = await fileResponse.json();
        name = data.name;
      }

      // Fetch Components
      const compResponse = await fetch(`https://api.figma.com/v1/files/${fileKey}/components`, {
        headers: { 'X-Figma-Token': token }
      });
      if (compResponse.ok) {
        const compData = await compResponse.json();
        if (compData.meta && compData.meta.components) {
             compData.meta.components.forEach((c: any) => {
                 components[c.name] = c.key;
             });
        }
      }

      // Fetch Styles
      const styleResponse = await fetch(`https://api.figma.com/v1/files/${fileKey}/styles`, {
        headers: { 'X-Figma-Token': token }
      });
      if (styleResponse.ok) {
        const styleData = await styleResponse.json();
        if (styleData.meta && styleData.meta.styles) {
             styleData.meta.styles.forEach((s: any) => {
                 styles[s.name] = s.key;
             });
        }
      }

    } catch (err) {
      console.error('Failed to fetch library details:', err);
      figma.notify('Failed to fetch full library details. Check your token.');
    }
  } else {
      figma.notify('No access token found. Please add one in settings to fetch library details.');
  }
  
  const newLib = {
    name: name,
    id: fileKey,
    key: fileKey,
    type: 'Remote',
    lastSynced: new Date().toISOString(),
    components: components,
    styles: styles
  };
  
  // Remove existing if present (update)
  const existingIndex = CONNECTED_LIBRARIES.findIndex(l => l.id === fileKey);
  if (existingIndex >= 0) {
      CONNECTED_LIBRARIES[existingIndex] = newLib;
      figma.notify(`Library "${name}" updated!`);
  } else {
      CONNECTED_LIBRARIES.push(newLib);
      figma.notify(`Library "${name}" added!`);
  }
  
  console.log('💾 Saving updated libraries list...');
  await saveConnectedLibraries();
  console.log('📤 Sending updated libraries to UI...');
  sendConnectedLibraries();
  updateMappingsFromConnected();
}

async function handleRemoveLibrary(libraryId: string) {
  CONNECTED_LIBRARIES = CONNECTED_LIBRARIES.filter(l => l.id !== libraryId);
  await saveConnectedLibraries();
  sendConnectedLibraries();
  figma.notify('Library removed.');
}

async function handleSaveAccessToken(token: string) {
  await figma.clientStorage.setAsync('figma_access_token', token);
  figma.notify('Access token saved securely.');
}

async function handleRemoveAccessToken() {
  await figma.clientStorage.deleteAsync('figma_access_token');
  figma.notify('Access token removed.');
}

async function handleGetAccessToken() {
  const token = await figma.clientStorage.getAsync('figma_access_token');
  figma.ui.postMessage({ type: 'ACCESS_TOKEN_LOADED', token: token || '' });
}

async function handleTestAccessToken() {
  const token = await figma.clientStorage.getAsync('figma_access_token');
  if (token) {
    figma.notify(`Success! Token found: ${token.substring(0, 4)}...`);
  } else {
    figma.notify('No access token found in storage.');
  }
}

// Call load on start
loadConnectedLibraries();

// Handle messages from the UI
figma.ui.onmessage = async (msg: any) => {
  switch (msg.type) {
    case 'SCAN_ALL':
      await handleScanFrames();
      break;
    case 'SWAP_LIBRARY':
      handleSwapLibrary('');
      break;
    case 'PERFORM_LIBRARY_SWAP':
      await performLibrarySwap(msg.components, msg.styles, msg.sourceLibrary, msg.targetLibrary);
      break;
    case 'GET_TARGET_COLOR':
      await getTargetColor(msg.tokenId, msg.styleName, msg.sourceLibrary, msg.targetLibrary);
      break;
    case 'UPDATE_LIBRARY_DEFINITIONS':
      await syncLibraryDefinitions();
      break;
    case 'add-library-by-link':
      await handleAddLibraryByLink(msg.link);
      break;
    case 'RESET_PLUGIN':
      await figma.clientStorage.setAsync('connected_libraries', []);
      CONNECTED_LIBRARIES = [];
      sendConnectedLibraries();
      figma.notify('Plugin data reset');
      await handleScanFrames();
      break;
    case 'ADD_LIBRARY_BY_KEY':
      await handleAddLibraryByKey(msg.fileKey);
      break;
    case 'REMOVE_LIBRARY':
      await handleRemoveLibrary(msg.libraryId);
      break;
    case 'GET_CONNECTED_LIBRARIES':
      console.log('📩 Received GET_CONNECTED_LIBRARIES request');
      sendConnectedLibraries();
      break;
    case 'SAVE_ACCESS_TOKEN':
      await handleSaveAccessToken(msg.token);
      break;
    case 'REMOVE_ACCESS_TOKEN':
      await handleRemoveAccessToken();
      break;
    case 'GET_ACCESS_TOKEN':
      await handleGetAccessToken();
      break;
    case 'TEST_ACCESS_TOKEN':
      await handleTestAccessToken();
      break;
    case 'SHOW_NATIVE_TOAST':
      figma.notify(msg.message || 'Swap completed successfully!');
      break;
    case 'close-plugin':
      figma.closePlugin();
      break;
    default:
      console.log('Unknown message type:', msg);
  }
};

// Scan selected frames for components and tokens
async function handleScanFrames(): Promise<void> {
  try {
    console.log('🔍 handleScanFrames called');
    const selection = figma.currentPage.selection;
    console.log('📦 Selection count:', selection.length);
    if (selection.length === 0) {
      figma.ui.postMessage({ type: 'SCAN_ALL_RESULT', ok: false, error: 'Please select at least one frame to scan' });
      return;
    }
    const components: ComponentInfo[] = [];
    const tokens: TokenInfo[] = [];
    for (const node of selection) {
      console.log('🔍 Checking node:', node.type, node.name);
      if (node.type === 'FRAME') {
        scannedFrame = node;  // Store the frame for later use
        console.log('✅ Found frame:', node.name);
        await scanNodeForAssets(node, components, tokens);
      }
    }
    console.log('📊 Scan complete. Found components:', components.length, 'tokens:', tokens.length);
    
    // Filter components and tokens to only include those from connected libraries
    const connectedLibraryNames = new Set(CONNECTED_LIBRARIES.map(l => l.name));
    
    // Also include libraries from JSONBin if they are considered "connected" or if we want to allow them
    // But user requirement is "ignore ... not included in connected libraries"
    // So we strictly filter by CONNECTED_LIBRARIES
    
    const filteredComponents = components.filter(c => connectedLibraryNames.has(c.library));
    const filteredTokens = tokens.filter(t => t.library && connectedLibraryNames.has(t.library));
    
    console.log('🔍 Filtered scan results:', {
        originalComponents: components.length,
        filteredComponents: filteredComponents.length,
        originalTokens: tokens.length,
        filteredTokens: filteredTokens.length
    });

    // Extract unique libraries and build library metadata with file IDs
    const libraryMap = new Map<string, { fileId?: string; components: number; tokens: number }>();
    
    // If no components or tokens found (after filtering), and we have a frame selected, check if we have any connected libraries
    console.log('📊 Checking for empty scan:', { 
        components: filteredComponents.length, 
        tokens: filteredTokens.length, 
        connectedLibsCount: CONNECTED_LIBRARIES.length,
        connectedLibsNames: CONNECTED_LIBRARIES.map(l => l.name)
    });
    
    if (filteredComponents.length === 0 && filteredTokens.length === 0) {
        if (CONNECTED_LIBRARIES.length === 0) {
            console.log('🚀 Sending SHOW_CONNECT_LIBRARY_VIEW');
            figma.ui.postMessage({ type: 'SHOW_CONNECT_LIBRARY_VIEW' });
            return;
        } else {
            console.log('⚠️ Scan empty, but libraries connected. Sending empty result.');
        }
    }

    // Add components to library map
    filteredComponents.forEach(c => {
      if (!libraryMap.has(c.library)) {
        libraryMap.set(c.library, { fileId: c.libraryFileId, components: 0, tokens: 0 });
      }
      const lib = libraryMap.get(c.library)!;
      lib.components += 1;
      if (c.libraryFileId && !lib.fileId) {
        lib.fileId = c.libraryFileId;
      }
    });
    
    // Add tokens to library map
    filteredTokens.forEach(t => {
      if (t.library && !libraryMap.has(t.library)) {
        libraryMap.set(t.library, { components: 0, tokens: 0 });
      }
      if (t.library) {
        libraryMap.get(t.library)!.tokens += 1;
      }
    });
    
    // Build libraries array with file IDs
    const libraries = Array.from(libraryMap.entries()).map(([libName, libInfo]) => ({
      name: libName,
      fileId: libInfo.fileId || null,
      thumbnail: LIBRARY_THUMBNAILS[libName] || null,
      componentCount: libInfo.components,
      tokenCount: libInfo.tokens
    }));
    
    figma.ui.postMessage({ type: 'SCAN_ALL_RESULT', ok: true, data: { components: filteredComponents, tokens: filteredTokens, libraries } });
  } catch (error) {
    console.error('❌ Scan error:', error);
    figma.ui.postMessage({ type: 'SCAN_ALL_RESULT', ok: false, error: `Scan failed: ${error instanceof Error ? error.message : 'Unknown error'}` });
  }
}

// Recursively scan nodes for components and design tokens
async function scanNodeForAssets(node: SceneNode, components: ComponentInfo[], tokens: TokenInfo[]): Promise<void> {
  if (Object.keys(COMPONENT_KEY_MAPPING).length === 0) {
      console.warn("⚠️ scanNodeForAssets: COMPONENT_KEY_MAPPING is empty! No components can be matched.");
  }

  if (node.type === 'INSTANCE') {
    try {
      const component = await node.getMainComponentAsync();
      if (!component) {
          // console.log(`⚠️ Instance ${node.name} has no main component.`);
      }
      if (component) {
        let foundLibrary: string | null = null;
        let foundName: string | null = null;
        let parentName: string | null = null;
        
        // Extract file ID from component key (format: fileId/pageId/componentId)
        const keyParts = component.key.split('/');
        const libraryFileId = keyParts[0] || undefined;
        const componentId = keyParts[keyParts.length - 1] || component.key; // Get the last part (component ID)
        
        console.log(`🔎 Inspecting component: ${component.name}`);
        console.log(`   - Key: ${component.key}`);
        console.log(`   - Remote: ${component.remote}`);
        console.log(`   - Extracted File ID: ${libraryFileId}`);
        
        // Try to match the component using the full key first, then try with extracted component ID
        let bestMatch = { library: null as string | null, name: null as string | null, score: 0, parentName: null as string | null };

        for (const libName of Object.keys(COMPONENT_KEY_MAPPING)) {
          for (const compName of Object.keys(COMPONENT_KEY_MAPPING[libName])) {
            const mappedKey = COMPONENT_KEY_MAPPING[libName][compName];
            
            // Debug logging for key matching
            // console.log(`🔍 Checking ${compName} in ${libName}: mapped=${mappedKey}, actual=${component.key}, id=${componentId}`);
            
            // Match scoring:
            // 5. Exact key match
            // 4. Component ID match (last part of key)
            // 3. Mapped key is contained in component key (for some remote library formats)
            // 2. Component key is contained in mapped key (reverse check)
            // 1. Name match (fallback) - CAUTION: This can be risky if names are not unique
            
            let score = 0;
            if (mappedKey === component.key) score = 5;
            else if (mappedKey === componentId) score = 4;
            else if (component.key.includes(mappedKey)) score = 3;
            else if (mappedKey.includes(componentId)) score = 2;
            else if (component.name === compName) score = 1;

            if (score > bestMatch.score) {
                let pName = null;
                if (component.parent && component.parent.type === 'COMPONENT_SET') {
                    pName = component.parent.name;
                } else {
                    pName = compName; // Use the mapped name as parent name if no component set
                }
                bestMatch = { library: libName, name: compName, score, parentName: pName };
            }
          }
        }

        // Fallback: If no match found in mapping, try to match by file ID directly
        if (!bestMatch.library) {
             // Check connected libraries for exact component key match
             console.log(`   - No mapping match. Checking ${CONNECTED_LIBRARIES.length} connected libraries...`);
             for (const lib of CONNECTED_LIBRARIES) {
                 if (lib.components) {
                     // Iterate entries to find key match
                     for (const [compName, compKey] of Object.entries(lib.components)) {
                         if (compKey === component.key) {
                             console.log(`✅ Found exact component key match in connected library: ${lib.name}`);
                             bestMatch.library = lib.name;
                             bestMatch.name = compName;
                             bestMatch.score = 5;
                             
                             // Try to determine parent name
                             if (component.parent && component.parent.type === 'COMPONENT_SET') {
                                 bestMatch.parentName = component.parent.name;
                             } else {
                                 bestMatch.parentName = compName;
                             }
                             break;
                         }
                     }
                 }
                 if (bestMatch.library) break;
             }
        }

        // Secondary Fallback: If still no match, try to match by file ID directly (less reliable)
        if (!bestMatch.library && libraryFileId) {
             const connectedLib = CONNECTED_LIBRARIES.find(lib => lib.id === libraryFileId || lib.key === libraryFileId);
             if (connectedLib) {
                 console.log(`⚠️ No component mapping found, but matched library by file ID: ${connectedLib.name}`);
                 bestMatch.library = connectedLib.name;
                 bestMatch.name = component.name;
                 bestMatch.score = 0.5; // Low score but enough to exist
                 
                 // Try to get a better parent name
                 if (component.parent && component.parent.type === 'COMPONENT_SET') {
                     bestMatch.parentName = component.parent.name;
                 } else {
                     bestMatch.parentName = component.name;
                 }
             } else {
                 // Last resort: Check if the file ID matches any connected library ID even partially
                 // This handles cases where the key format might be slightly different
                 const partialMatchLib = CONNECTED_LIBRARIES.find(lib => libraryFileId.includes(lib.id) || (lib.key && libraryFileId.includes(lib.key)));
                 if (partialMatchLib) {
                     console.log(`⚠️ Partial file ID match found: ${partialMatchLib.name}`);
                     bestMatch.library = partialMatchLib.name;
                     bestMatch.name = component.name;
                     bestMatch.score = 0.2;
                     bestMatch.parentName = component.name;
                 }
             }
        }
        
        if (bestMatch.library && bestMatch.name) {
          foundLibrary = bestMatch.library;
          foundName = bestMatch.name;
          parentName = bestMatch.parentName;
          console.log(`🏆 Best match for ${component.name}: ${foundLibrary} (Score: ${bestMatch.score})`);
        }
        if (foundLibrary && foundName) {
          const variantName = foundName; // The variant key from mapping
          const parentComponentName = parentName || foundName;
          console.log(`✅ Found component mapping: ${variantName} (parent: ${parentComponentName})`);
          if (!variantName.startsWith('.')) {
            components.push({ 
              id: component.id, 
              name: variantName,  // Variant name for swapping
              displayName: parentComponentName,  // Parent name for UI display
              library: foundLibrary, 
              remote: component.remote, 
              parentName: parentComponentName,
              libraryFileId: libraryFileId
            });
          }
        }
      }
    } catch (error) {}
  }
  await scanNodeForStyles(node, tokens);
  if ('children' in node && node.children) {
    for (const child of node.children) {
      await scanNodeForAssets(child, components, tokens);
    }
  }
}

// Scan a node for style tokens
async function scanNodeForStyles(node: SceneNode, tokens: TokenInfo[]): Promise<void> {
  console.log(`🔍 Scanning node: ${node.name} (type: ${node.type})`);
  
  // Check for paint styles (traditional Shark approach)
  if ('fillStyleId' in node && node.fillStyleId && typeof node.fillStyleId === 'string') {
    try {
      const paintStyle = await figma.getStyleByIdAsync(node.fillStyleId);
      if (paintStyle && paintStyle.type === 'PAINT') {
        console.log(`🎨 Found paint style: ${paintStyle.name}, key: ${paintStyle.key}`);
        // Determine library from style key
        let library = 'Unknown';
        for (const libName of Object.keys(STYLE_KEY_MAPPING)) {
          for (const styleName of Object.keys(STYLE_KEY_MAPPING[libName])) {
            console.log(`Checking style ${styleName} in ${libName}: ${STYLE_KEY_MAPPING[libName][styleName]} vs ${paintStyle.key}`);
            if (STYLE_KEY_MAPPING[libName][styleName] === paintStyle.key) {
              library = libName;
              console.log(`✅ Matched style key to library: ${library}`);
              break;
            }
          }
          if (library !== 'Unknown') break;
        }
        if (library === 'Unknown') {
          console.log(`❌ No match found for style key: ${paintStyle.key}`);
        }
        const tokenInfo: TokenInfo = { id: paintStyle.id, name: paintStyle.name, type: 'color', value: getColorValue(paintStyle.paints[0]), library };
        if (!tokens.find(t => t.id === tokenInfo.id)) tokens.push(tokenInfo);
      }
    } catch {}
  }
  
  // Check for variable-bound paints (Monkey approach)
  const nodeAny = node as any;
  if (nodeAny.boundVariables && nodeAny.boundVariables.fills && Array.isArray(nodeAny.boundVariables.fills)) {
    for (const boundVar of nodeAny.boundVariables.fills) {
      if (boundVar && typeof boundVar === 'object' && boundVar.type === 'VARIABLE_ALIAS') {
        try {
          const variable = await figma.variables.getVariableByIdAsync(boundVar.id);
          if (variable && variable.resolvedType === 'COLOR') {
            console.log(`🎨 Found variable-bound paint: ${variable.name} (${boundVar.id})`);
            
            // Determine library from variable key
            let library = 'Unknown';
            for (const libName of Object.keys(VARIABLE_KEY_MAPPING)) {
              for (const varName of Object.keys(VARIABLE_KEY_MAPPING[libName])) {
                if (VARIABLE_KEY_MAPPING[libName][varName] === variable.key) {
                  library = libName;
                  console.log(`✅ Matched variable key to library: ${library}`);
                  break;
                }
              }
              if (library !== 'Unknown') break;
            }
            
            // Fallback: check if this variable key matches any style key (handles cross-type tokens)
            if (library === 'Unknown') {
              for (const libName of Object.keys(STYLE_KEY_MAPPING)) {
                for (const styleName of Object.keys(STYLE_KEY_MAPPING[libName])) {
                  if (STYLE_KEY_MAPPING[libName][styleName] === variable.key) {
                    library = libName;
                    console.log(`✅ Matched variable key to style in library: ${library}`);
                    break;
                  }
                }
                if (library !== 'Unknown') break;
              }
            }
            
            // Get the color value from the variable
            const colorValue = variable.valuesByMode[Object.keys(variable.valuesByMode)[0]];
            let colorStr = '#000000';
            if (colorValue && typeof colorValue === 'object' && 'r' in colorValue) {
              colorStr = getColorValue({ type: 'SOLID', color: colorValue } as any);
            }
            
            const tokenInfo: TokenInfo = { id: variable.id, name: variable.name, type: 'color', value: colorStr, library };
            if (!tokens.find(t => t.id === tokenInfo.id)) tokens.push(tokenInfo);
          }
        } catch (err) {
          console.log(`⚠️ Failed to get variable: ${err}`);
        }
      }
    }
  }
  
  if ('textStyleId' in node && node.textStyleId && typeof node.textStyleId === 'string') {
    try {
      const textStyle = await figma.getStyleByIdAsync(node.textStyleId);
      if (textStyle && textStyle.type === 'TEXT') {
        console.log(`📝 Found text style: ${textStyle.name}, key: ${textStyle.key}`);
        // Determine library from style key
        let library = 'Unknown';
        for (const libName of Object.keys(STYLE_KEY_MAPPING)) {
          for (const styleName of Object.keys(STYLE_KEY_MAPPING[libName])) {
            if (STYLE_KEY_MAPPING[libName][styleName] === textStyle.key) {
              library = libName;
              console.log(`✅ Matched style key to library: ${library}`);
              break;
            }
          }
          if (library !== 'Unknown') break;
        }
        if (library === 'Unknown') {
          console.log(`❌ No match found for style key: ${textStyle.key}`);
        }
        const tokenInfo: TokenInfo = { id: textStyle.id, name: textStyle.name, type: 'typography', value: `${textStyle.fontSize}px ${textStyle.fontName?.family || 'Unknown'}`, library };
        if (!tokens.find(t => t.id === tokenInfo.id)) tokens.push(tokenInfo);
      }
    } catch {}
  }
  if ('effectStyleId' in node && node.effectStyleId && typeof node.effectStyleId === 'string') {
    try {
      const effectStyle = await figma.getStyleByIdAsync(node.effectStyleId);
      if (effectStyle && effectStyle.type === 'EFFECT') {
        console.log(`✨ Found effect style: ${effectStyle.name}, key: ${effectStyle.key}`);
        // Determine library from style key
        let library = 'Unknown';
        for (const libName of Object.keys(STYLE_KEY_MAPPING)) {
          for (const styleName of Object.keys(STYLE_KEY_MAPPING[libName])) {
            if (STYLE_KEY_MAPPING[libName][styleName] === effectStyle.key) {
              library = libName;
              console.log(`✅ Matched style key to library: ${library}`);
              break;
            }
          }
          if (library !== 'Unknown') break;
        }
        if (library === 'Unknown') {
          console.log(`❌ No match found for style key: ${effectStyle.key}`);
        }
        const tokenInfo: TokenInfo = { id: effectStyle.id, name: effectStyle.name, type: 'effect', value: 'Effect style', library };
        if (!tokens.find(t => t.id === tokenInfo.id)) tokens.push(tokenInfo);
      }
    } catch {}
  }
}

// Helper: extract color value from paint
function getColorValue(paint: Paint): string {
  if (paint.type === 'SOLID' && paint.color) {
    const r = Math.round(paint.color.r * 255);
    const g = Math.round(paint.color.g * 255);
    const b = Math.round(paint.color.b * 255);
    const hex = '#' + [r, g, b].map(x => x.toString(16).padStart(2, '0')).join('').toUpperCase();
    return hex;
  }
  return '#000000';
}

// Helper to normalize library names
function normalizeLibraryName(name: string): string {
  return name;
}

// Get target color for UI display
async function getTargetColor(tokenId: string, styleName: string, sourceLibrary: string, targetLibrary: string): Promise<void> {
  console.log(`🎨 Getting target color for: ${styleName}, from ${sourceLibrary} to ${targetLibrary}`);
  console.log(`📦 tokenId: ${tokenId}`);
  
  const normalizedTargetLibrary = normalizeLibraryName(targetLibrary);
  console.log(`📋 Normalized target library: ${normalizedTargetLibrary}`);
  
  // Check if target library uses variables instead of styles
  const targetVariableKey = VARIABLE_KEY_MAPPING[normalizedTargetLibrary]?.[styleName];
  console.log(`🔍 VARIABLE_KEY_MAPPING[${normalizedTargetLibrary}][${styleName}] = ${targetVariableKey}`);
  
  if (targetVariableKey) {
    console.log(`✅ Found variable binding for ${styleName}: ${targetVariableKey}`);
    try {
      // Import the variable to get its actual color value
      const variable = await figma.variables.importVariableByKeyAsync(targetVariableKey);
      console.log(`📥 Imported variable: ${variable?.name}, resolved type: ${variable?.resolvedType}`);
      if (variable && variable.resolvedType === 'COLOR') {
        // Get the color value from the variable
        const colorValue = variable.valuesByMode[Object.keys(variable.valuesByMode)[0]];
        console.log(`🎨 Variable color value object:`, colorValue);
        if (colorValue && typeof colorValue === 'object' && 'r' in colorValue) {
          const color = getColorValue({ type: 'SOLID', color: colorValue } as any);
          console.log(`🎨 Variable color resolved to: ${color}`);
          figma.ui.postMessage({ type: 'TARGET_COLOR_RESULT', tokenId, color });
          return;
        }
      }
    } catch (err) {
      console.warn(`❌ Failed to import variable ${styleName}:`, err);
    }
  }
  
  // Otherwise try to import as a style
  const targetStyleKey = STYLE_KEY_MAPPING[normalizedTargetLibrary]?.[styleName];
  
  console.log(`📍 Target library: ${normalizedTargetLibrary}, Style key: ${targetStyleKey}`);
  
  if (targetStyleKey) {
    try {
      const targetStyle = await figma.importStyleByKeyAsync(targetStyleKey);
      console.log(`✅ Imported style: ${targetStyle?.name}, type: ${targetStyle?.type}`);
      
      if (targetStyle && targetStyle.type === 'PAINT' && targetStyle.paints.length > 0) {
        const color = getColorValue(targetStyle.paints[0]);
        console.log(`🎨 Target color: ${color}`);
        figma.ui.postMessage({ type: 'TARGET_COLOR_RESULT', tokenId, color });
        return;
      }
    } catch (err) {
      console.warn(`❌ Failed to get target color for ${styleName}:`, err);
    }
  } else {
    console.warn(`❌ No style key found for ${styleName} in ${normalizedTargetLibrary}`);
  }
  
  // Fallback: return null if target not found
  console.log(`⚠️ Sending null color for tokenId: ${tokenId}`);
  figma.ui.postMessage({ type: 'TARGET_COLOR_RESULT', tokenId, color: null });
}

// Library swap stub
function handleSwapLibrary(libraryId: string): void {
  figma.ui.postMessage({ type: 'swap-complete', message: 'Library swap completed successfully' });
}

// Sync library definitions - outputs keys for local components and styles
async function syncLibraryDefinitions(): Promise<void> {
  console.log('🔄 Syncing library definitions...');
  
  const components: { name: string; key: string }[] = [];
  const styles: { name: string; key: string; type: string }[] = [];
  const variables: { name: string; id: string; key?: string }[] = [];
  
  // Load all pages first
  await figma.loadAllPagesAsync();
  
  // Scan ALL pages, not just current, to find components and variables
  let allNodes: SceneNode[] = [];
  for (const page of figma.root.children as PageNode[]) {
    console.log(`📄 Scanning page: ${page.name}`);
    const pageNodes = page.findAll();
    allNodes = allNodes.concat(pageNodes);
  }
  
  // Get all local components from all pages
  const localComponents = allNodes.filter(node => node.type === 'COMPONENT') as ComponentNode[];
  for (const comp of localComponents) {
    if (comp.key) {
      components.push({ name: comp.name, key: comp.key });
    }
  }
  
  // Get all local styles
  const localStyles = await figma.getLocalPaintStylesAsync();
  for (const style of localStyles) {
    if (style.key) {
      styles.push({ name: style.name, key: style.key, type: 'PAINT' });
    }
  }
  
  const localTextStyles = await figma.getLocalTextStylesAsync();
  for (const style of localTextStyles) {
    if (style.key) {
      styles.push({ name: style.name, key: style.key, type: 'TEXT' });
    }
  }
  
  const localEffectStyles = await figma.getLocalEffectStylesAsync();
  for (const style of localEffectStyles) {
    if (style.key) {
      styles.push({ name: style.name, key: style.key, type: 'EFFECT' });
    }
  }

  // Get all local variables using the correct API method
  try {
    const allVariables = await figma.variables.getLocalVariablesAsync();
    console.log(`📊 Found ${allVariables.length} local variables in file`);
    for (const variable of allVariables) {
      if (variable.name && variable.id) {
        variables.push({ name: variable.name, id: variable.id, key: variable.key });
        console.log(`  - ${variable.name}: ${variable.id} (key: ${variable.key})`);
      }
    }
  } catch (err) {
    console.warn(`⚠️ Could not get variables: ${err}`);
  }
  
  // Output to console for easy copying
  console.log('\n📦 COMPONENT KEYS:');
  components.forEach(comp => {
    console.log(`  '${comp.name}': '${comp.key}',`);
  });
  
  console.log('\n🎨 STYLE KEYS:');
  styles.forEach(style => {
    console.log(`  '${style.name}': '${style.key}', // ${style.type}`);
  });

  console.log('\n📦 VARIABLE IDs:');
  variables.forEach(variable => {
    console.log(`  '${variable.name}': '${variable.id}',`);
  });
  
  // Try to upload to JSONBin
  // await uploadToJSONBin(components, styles, variables);
  
  figma.notify(`Found ${components.length} components, ${styles.length} styles, and ${variables.length} variables. Check console for keys/IDs.`);
  figma.ui.postMessage({ 
    type: 'SYNC_COMPLETE', 
    components: components.length, 
    styles: styles.length,
    variables: variables.length
  });
}

// Helper: Recursively find all instances by name and library
async function findInstancesByNameAsync(node: SceneNode, name: string, library: string): Promise<InstanceNode[]> {
  let found: InstanceNode[] = [];
  if (node.type === 'INSTANCE') {
    try {
      const mainComponent = await node.getMainComponentAsync();
      if (mainComponent) {
        console.log(`🔍 Found instance: ${mainComponent.name}, key: ${mainComponent.key}, looking for: ${name} in ${library}`);
        // Check if the main component's key matches the mapping for this component name
        const expectedKey = COMPONENT_KEY_MAPPING[library]?.[name];
        if (expectedKey && mainComponent.key === expectedKey) {
          console.log(`✅ Match found! Adding instance.`);
          found.push(node);
        }
      }
    } catch (err) {
      // Ignore errors for nodes that can't resolve mainComponent
    }
  }
  if ('children' in node && node.children) {
    for (const child of node.children) {
      const childFound = await findInstancesByNameAsync(child, name, library);
      found = found.concat(childFound);
    }
  }
  return found;
}

// Recursively swap styles in a node and its children
async function swapStylesInNode(node: SceneNode, styleName: string, sourceLibrary: string, targetLibrary: string): Promise<number> {
  let swapCount = 0;
  
  // Normalize library names for legacy lookups
  const normalizedSourceLibrary = normalizeLibraryName(sourceLibrary);
  const normalizedTargetLibrary = normalizeLibraryName(targetLibrary);

  // Find source library metadata
  const sourceLibObj = CONNECTED_LIBRARIES.find(l => l.name === sourceLibrary);
  let sourceStyleKey: string | undefined;
  
  if (sourceLibObj && sourceLibObj.styles) {
      sourceStyleKey = sourceLibObj.styles[styleName];
  }
  
  // Fallback to legacy mapping
  if (!sourceStyleKey) {
      sourceStyleKey = STYLE_KEY_MAPPING[sourceLibrary]?.[styleName];
  }
  
  // Fallback: check if source is a variable in the source library
  if (!sourceStyleKey) {
    sourceStyleKey = VARIABLE_KEY_MAPPING[sourceLibrary]?.[styleName];
    if (sourceStyleKey) {
      console.log(`  📌 Found source as variable in ${sourceLibrary}: ${styleName} -> ${sourceStyleKey}`);
    }
  }
  
  if (!sourceStyleKey) {
    console.warn(`⚠️ Source style/variable mapping not found for '${styleName}' in ${sourceLibrary}`);
    return 0;
  }

  // Find target library metadata
  const targetLibObj = CONNECTED_LIBRARIES.find(l => l.name === targetLibrary);
  let targetStyleKey: string | undefined;
  
  if (targetLibObj && targetLibObj.styles) {
      targetStyleKey = targetLibObj.styles[styleName];
  }
  
  // Fallback to legacy mapping
  if (!targetStyleKey) {
      targetStyleKey = STYLE_KEY_MAPPING[targetLibrary]?.[styleName];
  }

  // Try to get the target variable KEY for this style name (Legacy fallback)
  const targetVariableKey = VARIABLE_KEY_MAPPING[targetLibrary]?.[styleName];
  const targetVariableId = VARIABLE_ID_MAPPING[targetLibrary]?.[styleName];

  console.log(`  📋 Processing node: ${node.name} (type: ${node.type}) looking for style: ${styleName}`);

  // Process fill styles on any node that has them (including FRAME)
  if ('fillStyleId' in node && node.fillStyleId && typeof node.fillStyleId === 'string') {
    try {
      const currentStyle = await figma.getStyleByIdAsync(node.fillStyleId);
      if (currentStyle && currentStyle.key === sourceStyleKey) {
        console.log(`  ✅ Found ${styleName} style on ${node.type}`);
        
        // If we have a target style key from dynamic metadata, use it
        if (targetStyleKey) {
             try {
                 const importedStyle = await figma.importStyleByKeyAsync(targetStyleKey);
                 if (importedStyle) {
                     (node as any).fillStyleId = importedStyle.id;
                     swapCount++;
                     console.log(`  ✅ Swapped style to ${targetLibrary}/${styleName}`);
                     return swapCount; // Return early if swapped
                 }
             } catch (e) {
                 console.warn(`  ⚠️ Failed to import target style: ${e}`);
             }
        }

        // If we have a target variable KEY or ID, bind to it (Legacy/Variable logic)
        if (targetVariableKey || targetVariableId) {
          try {
            console.log(`  🔄 Binding to Monkey variable: ${targetVariableKey || targetVariableId}`);
            const nodeAny = node as any;
            
            // Try setting the fill with variable binding directly
            if (Array.isArray(nodeAny.fills) && nodeAny.fills.length > 0) {
              try {
                console.log(`  🔄 Attempting to bind variable: ${styleName} (KEY: ${targetVariableKey}, ID: ${targetVariableId})`);
                
                let variable: Variable | null = null;
                
                // First, try to import by KEY (most reliable for library variables)
                if (targetVariableKey) {
                  try {
                    console.log(`  🔍 Attempting importVariableByKeyAsync("${targetVariableKey}")...`);
                    variable = await figma.variables.importVariableByKeyAsync(targetVariableKey);
                    if (variable) {
                      console.log(`  ✅ Found variable by KEY: ${variable.name}`);
                    } else {
                      console.log(`  ⚠️ importVariableByKeyAsync returned null`);
                    }
                  } catch (keyErr) {
                    console.log(`  ⚠️ importVariableByKeyAsync failed: ${keyErr}`);
                  }
                }
                
                // Fallback: search by name in local variables (includes imported library variables)
                if (!variable) {
                  console.log(`  🔍 Falling back to name search: "${styleName}"...`);
                  try {
                    const localVars = await figma.variables.getLocalVariablesAsync();
                    console.log(`  📊 Searching ${localVars.length} local variables (includes imported library variables)`);
                    console.log(`  📋 Available names: ${localVars.map(v => v.name).join(', ')}`);
                    
                    variable = localVars.find(v => v.name === styleName) || null;
                    if (variable) {
                      console.log(`  ✅ Found variable by name: ${variable.name}`);
                    } else {
                      console.log(`  ⚠️ No variable named '${styleName}' found in local or library variables`);
                    }
                  } catch (searchErr) {
                    console.log(`  ❌ Name search failed: ${searchErr}`);
                  }
                }
                
                if (variable) {
                  console.log(`  📌 Using variable: ${variable.name} (id: ${variable.id}, type: ${variable.resolvedType})`);
                  // Use the official API to bind the variable to paint
                  const boundPaint = figma.variables.setBoundVariableForPaint(
                    { type: 'SOLID', color: { r: 0, g: 0, b: 0 } },  // Base paint (color doesn't matter)
                    'color',                                          // Field to bind
                    variable                                          // The variable to use
                  );
                  
                  console.log(`  🎨 Bound paint created, assigning to fills...`);
                  nodeAny.fills = [boundPaint];
                  console.log(`  ✅ Swapped fill to variable '${styleName}' on ${node.type}`);
                  swapCount++;
                } else {
                  console.log(`  ❌ Could not find variable '${styleName}' (KEY: ${targetVariableKey}, ID: ${targetVariableId})`);
                }
              } catch (bindErr) {
                console.log(`  ❌ Error binding fill variable: ${bindErr}`);
              }
            }
          } catch (e) {
            console.log(`  ℹ️ Could not bind variable directly, trying style swap: ${e}`);
            // Fallback: try to swap to a style in the target library
            const targetStyleKey = STYLE_KEY_MAPPING[normalizedTargetLibrary]?.[styleName];
            if (targetStyleKey) {
              try {
                const targetStyle = await figma.importStyleByKeyAsync(targetStyleKey);
                if (targetStyle && targetStyle.type === 'PAINT') {
                  // Use async method for safer clearing
                  if (typeof (node as any).setFillStyleIdAsync === 'function') {
                    await (node as any).setFillStyleIdAsync(targetStyle.id);
                    console.log(`  ✅ Swapped fill to Monkey style '${styleName}' on ${node.type}`);
                    swapCount++;
                  }
                }
              } catch (importErr) {
                console.log(`  ❌ Could not import target style: ${importErr}`);
              }
            }
          }
        } else {
          console.log(`  ⚠️ No variable ID found for '${styleName}' in ${normalizedTargetLibrary}`);
        }
      }
    } catch (styleErr) {
      // Node doesn't have a valid fill style, skip it
    }
  }
  
  // Process variable-bound fills (swapping FROM Monkey variables TO Shark styles)
  const nodeAny = node as any;
  if (nodeAny.boundVariables && nodeAny.boundVariables.fills && Array.isArray(nodeAny.boundVariables.fills)) {
    try {
      // Check if any fills are bound to variables matching the source library
      for (const boundVar of nodeAny.boundVariables.fills) {
        if (boundVar && typeof boundVar === 'object' && boundVar.type === 'VARIABLE_ALIAS') {
          try {
            const variable = await figma.variables.getVariableByIdAsync(boundVar.id);
            if (variable && variable.resolvedType === 'COLOR' && variable.name === styleName) {
              // Check if this variable belongs to the source library
              let isFromSourceLibrary = false;
              for (const varName of Object.keys(VARIABLE_KEY_MAPPING[sourceLibrary] || {})) {
                if (varName === styleName) {
                  isFromSourceLibrary = true;
                  break;
                }
              }
              
              if (isFromSourceLibrary) {
                console.log(`  ✅ Found variable-bound fill: ${styleName} on ${node.type}`);
                
                // Try to swap to target style
                const targetStyleKey = STYLE_KEY_MAPPING[normalizedTargetLibrary]?.[styleName];
                if (targetStyleKey) {
                  try {
                    const targetStyle = await figma.importStyleByKeyAsync(targetStyleKey);
                    if (targetStyle && targetStyle.type === 'PAINT') {
                      if (typeof nodeAny.setFillStyleIdAsync === 'function') {
                        await nodeAny.setFillStyleIdAsync(targetStyle.id);
                        console.log(`  ✅ Swapped variable-bound fill to Shark style '${styleName}' on ${node.type}`);
                        swapCount++;
                      }
                    }
                  } catch (styleImportErr) {
                    console.log(`  ❌ Could not import target style: ${styleImportErr}`);
                  }
                }
              }
            }
          } catch (varErr) {
            // Skip if variable can't be fetched
          }
        }
      }
    } catch (boundVarErr) {
      // Skip if boundVariables can't be accessed
    }
  }

  // Process stroke styles on any node that has them
  if ('strokeStyleId' in node && node.strokeStyleId && typeof node.strokeStyleId === 'string') {
    try {
      const currentStyle = await figma.getStyleByIdAsync(node.strokeStyleId);
      if (currentStyle && currentStyle.key === sourceStyleKey) {
        console.log(`  ✅ Found ${styleName} stroke on ${node.type}`);
        
        if (targetVariableKey || targetVariableId) {
          try {
            console.log(`  🔄 Binding stroke to Monkey variable: ${targetVariableKey || targetVariableId}`);
            const nodeAny = node as any;
            
            if (Array.isArray(nodeAny.strokes) && nodeAny.strokes.length > 0) {
              try {
                console.log(`  🔄 Attempting to bind variable: ${styleName} (KEY: ${targetVariableKey}, ID: ${targetVariableId})`);
                
                let variable: Variable | null = null;
                
                // First, try to import by KEY (most reliable for library variables)
                if (targetVariableKey) {
                  try {
                    console.log(`  🔍 Attempting importVariableByKeyAsync("${targetVariableKey}")...`);
                    variable = await figma.variables.importVariableByKeyAsync(targetVariableKey);
                    if (variable) {
                      console.log(`  ✅ Found variable by KEY: ${variable.name}`);
                    } else {
                      console.log(`  ⚠️ importVariableByKeyAsync returned null`);
                    }
                  } catch (keyErr) {
                    console.log(`  ⚠️ importVariableByKeyAsync failed: ${keyErr}`);
                  }
                }
                
                // Fallback: search by name in local variables (includes imported library variables)
                if (!variable) {
                  console.log(`  🔍 Falling back to name search: "${styleName}"...`);
                  try {
                    const localVars = await figma.variables.getLocalVariablesAsync();
                    console.log(`  📊 Searching ${localVars.length} local variables (includes imported library variables)`);
                    console.log(`  📋 Available names: ${localVars.map(v => v.name).join(', ')}`);
                    
                    variable = localVars.find(v => v.name === styleName) || null;
                    if (variable) {
                      console.log(`  ✅ Found variable by name: ${variable.name}`);
                    } else {
                      console.log(`  ⚠️ No variable named '${styleName}' found in local or library variables`);
                    }
                  } catch (searchErr) {
                    console.log(`  ❌ Name search failed: ${searchErr}`);
                  }
                }
                
                if (variable) {
                  console.log(`  📌 Using variable: ${variable.name} (id: ${variable.id}, type: ${variable.resolvedType})`);
                  // Use the official API to bind the variable to paint
                  const boundPaint = figma.variables.setBoundVariableForPaint(
                    { type: 'SOLID', color: { r: 0, g: 0, b: 0 } },  // Base paint (color doesn't matter)
                    'color',                                          // Field to bind
                    variable                                          // The variable to use
                  );
                  
                  console.log(`  🎨 Bound paint created, assigning to strokes...`);
                  nodeAny.strokes = [boundPaint];
                  console.log(`  ✅ Swapped stroke to variable '${styleName}' on ${node.type}`);
                  swapCount++;
                } else {
                  console.log(`  ❌ Could not find variable '${styleName}' (KEY: ${targetVariableKey}, ID: ${targetVariableId})`);
                }
              } catch (bindErr) {
                console.log(`  ❌ Error binding stroke variable: ${bindErr}`);
              }
            }
          } catch (e) {
            console.log(`  ⚠️ Could not bind stroke variable: ${e}`);
          }
        }
      }
    } catch (styleErr) {
      // Node doesn't have a valid stroke style, skip it
    }
  }

  // Process text styles on TEXT nodes
  if (node.type === 'TEXT' && 'textStyleId' in node && node.textStyleId && typeof node.textStyleId === 'string') {
    try {
      const currentStyle = await figma.getStyleByIdAsync(node.textStyleId);
      if (currentStyle && currentStyle.type === 'TEXT' && currentStyle.key === sourceStyleKey) {
        console.log(`  ✅ Found ${styleName} text style on ${node.type}`);
        
        // Get the target text style and apply it
        const targetStyleKey = STYLE_KEY_MAPPING[normalizedTargetLibrary]?.[styleName];
        if (targetStyleKey) {
          try {
            const targetStyle = await figma.importStyleByKeyAsync(targetStyleKey);
            if (targetStyle && targetStyle.type === 'TEXT') {
              const textNode = node as TextNode;
              // Use async method for text style assignment
              await textNode.setTextStyleIdAsync(targetStyle.id);
              console.log(`  ✅ Swapped text style to '${styleName}' on TEXT node`);
              swapCount++;
            }
          } catch (importErr) {
            console.log(`  ❌ Could not import target text style: ${importErr}`);
          }
        }
      }
    } catch (styleErr) {
      // Node doesn't have a valid text style, skip it
    }
  }

  // Recursively process children
  if ('children' in node && node.children) {
    for (const child of node.children) {
      swapCount += await swapStylesInNode(child, styleName, sourceLibrary, targetLibrary);
    }
  }

  return swapCount;
}

// Apply target library styles to instance shapes by matching style names
function restoreTargetStyles(instanceNode: SceneNode, targetComponent: ComponentNode, targetLibrary: string): void {
  try {
    console.log('🎨 Applying target library styles to instance');
    
    // Map of shape names to their style names in target
    const targetStyleNames = new Map<string, string>();
    
    function collectTargetStyleNames(node: SceneNode): void {
      if ((node.type === 'RECTANGLE' || node.type === 'ELLIPSE' || node.type === 'POLYGON' || node.type === 'STAR')) {
        const inst = node as any;
        if (inst.fillStyleId) {
          try {
            const style = figma.getStyleById(inst.fillStyleId);
            if (style) {
              console.log(`📍 Target shape "${node.name}" uses style: "${style.name}"`);
              targetStyleNames.set(node.name, style.name);
            }
          } catch (e) {
            console.warn(`Could not get style for ${node.name}:`, e);
          }
        }
      }
      
      if ('children' in node) {
        for (const child of node.children) {
          collectTargetStyleNames(child);
        }
      }
    }
    
    // Collect style names from target component
    collectTargetStyleNames(targetComponent);
    console.log('Target style names:', Array.from(targetStyleNames.entries()));
    
    // Apply styles to instance shapes by name
    function applyTargetStyles(node: SceneNode): void {
      if ((node.type === 'RECTANGLE' || node.type === 'ELLIPSE' || node.type === 'POLYGON' || node.type === 'STAR')) {
        const styleName = targetStyleNames.get(node.name);
        if (styleName) {
          try {
            const inst = node as any;
            
            // Search for the style in the imported file/library
            // The style should be available through the file key mapping
            console.log(`🔍 Looking for style "${styleName}" in target library...`);
            
            // Try to find the style by name in all available styles
            // We'll iterate through all library files and search
            const allLibraryFiles = figma.getSharedLibraryFiles();
            
            for (const libFile of allLibraryFiles) {
              console.log(`📚 Checking library file: ${libFile.name}`);
              
              // Get all paint styles from this library
              try {
                const paintStyles = libFile.getSharedPluginData('figma', 'paintStyles');
                if (paintStyles) {
                  const styles = JSON.parse(paintStyles);
                  console.log(`  Found styles:`, styles);
                }
              } catch (e) {
                console.warn(`  Could not get styles from ${libFile.name}:`, e);
              }
            }
            
            // Also try getLocalPaintStylesAsync to search for the style
            const allPaintStyles = figma.getLocalPaintStyles();
            console.log(`🔎 Searching in ${allPaintStyles.length} local styles for "${styleName}"`);
            
            for (const style of allPaintStyles) {
              console.log(`  Style: ${style.name} (id: ${style.id})`);
              if (style.name === styleName) {
                console.log(`✅ Found matching style! Applying to "${node.name}"`);
                inst.fillStyleId = style.id;
                return;
              }
            }
            
            console.log(`⚠️ Style "${styleName}" not found in document`);
          } catch (e) {
            console.warn(`Error applying style to ${node.name}:`, e);
          }
        }
      }
      
      if ('children' in node) {
        for (const child of node.children) {
          applyTargetStyles(child);
        }
      }
    }
    
    // Apply styles to instance
    applyTargetStyles(instanceNode);
  } catch (error) {
    console.warn('Error applying target styles:', error);
  }
}

// Add detailed swap logic for PERFORM_LIBRARY_SWAP
async function performLibrarySwap(components: any[], styles: any[], sourceLibrary: string, targetLibrary: string) {
  console.log('🔄 performLibrarySwap called!');
  console.log('Components:', components);
  console.log('Styles:', styles);
  console.log('Source Library:', sourceLibrary);
  console.log('Target Library:', targetLibrary);

  // Find target library in connected libraries
  const targetLibObj = CONNECTED_LIBRARIES.find(l => l.name === targetLibrary);
  if (!targetLibObj) {
      console.warn(`Target library "${targetLibrary}" not found in connected libraries.`);
  }

  let swapCount = 0;
  let errorCount = 0;
  let errorDetails: string[] = [];
  let totalInstancesFound = 0;

  for (const comp of components) {
    try {
      console.log(`🔄 Processing component: ${comp.name}`);
      // Use the scanned frame if available, otherwise fall back to selection
      const nodesToProcess = scannedFrame ? [scannedFrame] : figma.currentPage.selection;
      
      let foundInstancesForThisComponent = false;
      for (const node of nodesToProcess) {
        if (node.type === 'FRAME' || node.type === 'GROUP') {
          const instances = await findInstancesByNameAsync(node, comp.name, sourceLibrary);
          totalInstancesFound += instances.length;
          if (instances.length === 0) {
            continue;  // Don't count as error yet, check other nodes
          }
          foundInstancesForThisComponent = true;
          for (const instance of instances) {
            // Look up target key in connected library metadata
            let targetKey: string | undefined;
            if (targetLibObj && targetLibObj.components) {
                targetKey = targetLibObj.components[comp.name];
            }
            
            // Fallback to legacy mapping if not found in dynamic metadata
            if (!targetKey) {
                 // Try legacy mapping if it exists
                 if (COMPONENT_KEY_MAPPING[targetLibrary] && COMPONENT_KEY_MAPPING[targetLibrary][comp.name]) {
                     targetKey = COMPONENT_KEY_MAPPING[targetLibrary][comp.name];
                 }
            }

            if (!targetKey) {
              errorDetails.push(`No mapping for component '${comp.name}' in target library '${targetLibrary}'`);
              errorCount++;
              continue;
            }
            try {
              const importedComponent = await figma.importComponentByKeyAsync(targetKey);
              console.log(`  ✅ Imported target component`);
              
              // Don't try to access componentProperties on the component definition
              // Properties will be accessible on instances after the swap
              
              // Store position only (not size, which becomes read-only after swap)
              const instanceNode = instance as InstanceNode;
              const x = instanceNode.x;
              const y = instanceNode.y;
              const rotation = instanceNode.rotation;
              
              // Capture component property overrides and layout from nested instances
              const nestedInstanceOverrides = new Map<number, Record<string, any>>();
              const nestedInstanceLayouts = new Map<number, Record<string, any>>();
              let nestedInstanceIndex = 0;
              async function captureNestedOverrides(node: SceneNode) {
                if (node.type === 'INSTANCE') {
                  const instNode = node as InstanceNode;
                  const instAny = instNode as any;
                  
                  // Capture component property overrides
                  if (instAny.componentProperties && typeof instAny.componentProperties === 'object') {
                    const overrides: Record<string, any> = {};
                    for (const [propName, propValue] of Object.entries(instAny.componentProperties)) {
                      if (propValue && typeof propValue === 'object' && 'value' in propValue) {
                        const cleanPropName = propName.split('#')[0];
                        overrides[cleanPropName] = (propValue as any).value;
                      }
                    }
                    if (Object.keys(overrides).length > 0) {
                      nestedInstanceOverrides.set(nestedInstanceIndex, overrides);
                      console.log(`  ⚙️ Captured overrides for nested instance #${nestedInstanceIndex}: ${Object.keys(overrides).join(', ')}`);
                    }
                  }
                  
                  // Capture layout overrides
                  const layoutProps: Record<string, any> = {};
                  if (instNode.layoutAlign !== 'STRETCH') layoutProps.layoutAlign = instNode.layoutAlign;
                  if (instNode.layoutGrow !== 0) layoutProps.layoutGrow = instNode.layoutGrow;
                  if (instNode.layoutMode !== 'NONE') layoutProps.layoutMode = instNode.layoutMode;
                  if (instNode.layoutPositioning !== 'AUTO') layoutProps.layoutPositioning = instNode.layoutPositioning;
                  if (instNode.layoutWrap !== 'NO_WRAP') layoutProps.layoutWrap = instNode.layoutWrap;
                  if ('maxHeight' in instNode && instNode.maxHeight !== null) layoutProps.maxHeight = instNode.maxHeight;
                  if ('maxWidth' in instNode && instNode.maxWidth !== null) layoutProps.maxWidth = instNode.maxWidth;
                  if ('minHeight' in instNode && instNode.minHeight !== null) layoutProps.minHeight = instNode.minHeight;
                  if ('minWidth' in instNode && instNode.minWidth !== null) layoutProps.minWidth = instNode.minWidth;
                  if (instNode.paddingBottom !== 0) layoutProps.paddingBottom = instNode.paddingBottom;
                  if (instNode.paddingLeft !== 0) layoutProps.paddingLeft = instNode.paddingLeft;
                  if (instNode.paddingRight !== 0) layoutProps.paddingRight = instNode.paddingRight;
                  if (instNode.paddingTop !== 0) layoutProps.paddingTop = instNode.paddingTop;
                  if (instNode.primaryAxisAlignItems !== 'MIN') layoutProps.primaryAxisAlignItems = instNode.primaryAxisAlignItems;
                  if (instNode.primaryAxisSizingMode !== 'AUTO') layoutProps.primaryAxisSizingMode = instNode.primaryAxisSizingMode;
                  if (instNode.secondaryAxisAlignItems !== 'MIN') layoutProps.secondaryAxisAlignItems = instNode.secondaryAxisAlignItems;
                  if (instNode.secondaryAxisSizingMode !== 'AUTO') layoutProps.secondaryAxisSizingMode = instNode.secondaryAxisSizingMode;
                  // Always capture width and height for resizing
                  layoutProps.width = instNode.width;
                  layoutProps.height = instNode.height;
                  if (Object.keys(layoutProps).length > 0) {
                    nestedInstanceLayouts.set(nestedInstanceIndex, layoutProps);
                    console.log(`  📐 Captured layout for nested instance #${nestedInstanceIndex}: ${Object.keys(layoutProps).join(', ')}`);
                  }
                  
                  nestedInstanceIndex++;
                }
                if ('children' in node) {
                  for (const child of node.children) {
                    await captureNestedOverrides(child);
                  }
                }
              }
              console.log('🔄 Capturing nested instance overrides before swap...');
              await captureNestedOverrides(instanceNode);
              
              // Capture text values from all text nodes in the instance before swap
              // Use unique keys based on index to handle duplicate names
              const textValues = new Map<string, { text: string; fontName: any }>();
              let textNodeIndex = 0;
              async function captureTextValues(node: SceneNode, path: string = '') {
                const nodePath = path ? `${path}/${node.name}` : node.name;
                if (node.type === 'TEXT') {
                  const textNode = node as TextNode;
                  // Create unique key for each text node (index-based)
                  const uniqueKey = `text_${textNodeIndex++}`;
                  textValues.set(uniqueKey, {
                    text: textNode.characters,
                    fontName: textNode.fontName
                  });
                  console.log(`  📝 Captured text #${textNodeIndex - 1} from "${nodePath}": "${textNode.characters.substring(0, 30)}..."`);
                }
                if ('children' in node) {
                  for (const child of node.children) {
                    await captureTextValues(child, nodePath);
                  }
                }
              }
              console.log('🔄 Capturing text values before swap...');
              await captureTextValues(instanceNode);
              console.log(`  Total text nodes captured: ${textValues.size}`);
              
              // Perform the swap
              instance.swapComponent(importedComponent);
              console.log(`✅ Swapped component: ${instance.name}`);
              
              // Remove all overrides so instance uses only target library defaults
              try {
                const inst = instanceNode as any;
                if (typeof inst.removeOverrides === 'function') {
                  inst.removeOverrides();
                  console.log('✅ Removed all overrides');
                } else if (typeof inst.resetAllOverrides === 'function') {
                  // Fallback to resetAllOverrides if removeOverrides doesn't exist
                  inst.resetAllOverrides();
                  console.log('✅ Reset all overrides (fallback)');
                }
              } catch (e) {
                console.warn('Error removing overrides:', e);
              }
              
              // Reapply captured text values using same index-based approach
              console.log('🔄 Reapplying text values after swap...');
              let reapplyIndex = 0;
              async function reapplyTextValues(node: SceneNode) {
                if (node.type === 'TEXT') {
                  const uniqueKey = `text_${reapplyIndex++}`;
                  const captured = textValues.get(uniqueKey);
                  if (captured) {
                    try {
                      const textNode = node as TextNode;
                      // Load the font before setting text
                      if (captured.fontName && typeof captured.fontName === 'object') {
                        await figma.loadFontAsync(captured.fontName);
                      }
                      textNode.characters = captured.text;
                      console.log(`  ✅ Reapplied text #${reapplyIndex - 1} to "${node.name}": "${captured.text.substring(0, 30)}..."`);
                    } catch (e) {
                      console.warn(`  ❌ Could not reapply text #${reapplyIndex - 1}:`, e);
                    }
                  }
                }
                if ('children' in node) {
                  for (const child of node.children) {
                    await reapplyTextValues(child);
                  }
                }
              }
              await reapplyTextValues(instanceNode);
              console.log('✅ Text reapplication complete');
              
              // Reapply captured nested instance property overrides
              if (nestedInstanceOverrides.size > 0) {
                console.log(`🔄 Reapplying nested instance overrides after swap...`);
                let reapplyOverrideIndex = 0;
                async function reapplyOverrides(node: SceneNode) {
                  if (node.type === 'INSTANCE') {
                    const capturedOverrides = nestedInstanceOverrides.get(reapplyOverrideIndex);
                    
                    if (capturedOverrides && Object.keys(capturedOverrides).length > 0) {
                      try {
                        // Get property definitions from the swapped target component
                        const instNode = node as InstanceNode;
                        const instAny = instNode as any;
                        const propsByBase: {[base: string]: string} = {};
                        
                        if (instAny.componentProperties) {
                          for (const propName of Object.keys(instAny.componentProperties)) {
                            const base = propName.split('#')[0];
                            propsByBase[base] = propName;
                          }
                        }
                        
                        // Build new properties using target's full property names
                        const newProps: {[key: string]: any} = {};
                        for (const [baseName, value] of Object.entries(capturedOverrides)) {
                          const fullPropName = propsByBase[baseName];
                          if (fullPropName) {
                            newProps[fullPropName] = value;
                          }
                        }
                        
                        if (Object.keys(newProps).length > 0) {
                          instNode.setProperties(newProps);
                          console.log(`  ✅ Applied ${Object.keys(newProps).length} properties to nested instance #${reapplyOverrideIndex}`);
                        }
                      } catch (e) {
                        console.warn(`  ❌ Could not reapply overrides for nested instance #${reapplyOverrideIndex}:`, e);
                      }
                    }
                    reapplyOverrideIndex++;
                  }
                  if ('children' in node) {
                    for (const child of node.children) {
                      await reapplyOverrides(child);
                    }
                  }
                }
                await reapplyOverrides(instanceNode);
                console.log('✅ Nested instance override reapplication complete');
              }
              
              // Reapply captured layout overrides
              if (nestedInstanceLayouts.size > 0) {
                console.log('🔄 Reapplying layout overrides after swap...');
                let reapplyLayoutIndex = 0;
                async function reapplyLayouts(node: SceneNode) {
                  if (node.type === 'INSTANCE') {
                    const capturedLayout = nestedInstanceLayouts.get(reapplyLayoutIndex);
                    if (capturedLayout && Object.keys(capturedLayout).length > 0) {
                      try {
                        const instNode = node as InstanceNode;
                        // Only set properties that are actually settable on nodes
                        const writableProps = ['layoutAlign', 'layoutGrow', 'layoutMode', 'layoutPositioning', 'layoutWrap', 
                                             'paddingBottom', 'paddingLeft', 'paddingRight', 'paddingTop'];
                        let applied = 0;
                        for (const [key, value] of Object.entries(capturedLayout)) {
                          if (writableProps.includes(key)) {
                            try {
                              (instNode as any)[key] = value;
                              applied++;
                            } catch (e) {
                              // Skip properties that can't be set
                            }
                          }
                        }
                        
                        // Handle width and height separately using resizeWithoutConstraints
                        if ('width' in capturedLayout && 'height' in capturedLayout) {
                          try {
                            const instNodeAny = instNode as any;
                            instNodeAny.resizeWithoutConstraints(capturedLayout.width, capturedLayout.height);
                            applied++;
                          } catch (e) {
                            // Fallback: try setting width and height directly
                            try {
                              instNode.resize(capturedLayout.width, capturedLayout.height);
                              applied++;
                            } catch (e2) {
                              // Skip if both methods fail
                            }
                          }
                        }
                        
                        if (applied > 0) {
                          console.log(`  ✅ Applied ${applied} layout properties to nested instance #${reapplyLayoutIndex}`);
                        }
                      } catch (e) {
                        console.warn(`  ❌ Could not reapply layout for nested instance #${reapplyLayoutIndex}:`, e);
                      }
                    }
                    reapplyLayoutIndex++;
                  }
                  if ('children' in node) {
                    for (const child of node.children) {
                      await reapplyLayouts(child);
                    }
                  }
                }
                await reapplyLayouts(instanceNode);
                console.log('✅ Layout override reapplication complete');
              }
              
              // Restore position only
              instanceNode.x = x;
              instanceNode.y = y;
              instanceNode.rotation = rotation;
              
              swapCount++;
            } catch (swapErr) {
              errorDetails.push(`Failed to swap '${comp.name}': ${swapErr instanceof Error ? swapErr.message : swapErr}`);
              errorCount++;
            }
          }
        }
      }
      // Only report error if no instances were found for this component at all AND nothing has been swapped yet
      if (!foundInstancesForThisComponent && totalInstancesFound === 0) {
        errorDetails.push(`No matching instances found for component '${comp.name}' in selection.`);
        errorCount++;
      }
    } catch (err) {
      errorDetails.push(`Error processing component '${comp.name}': ${err instanceof Error ? err.message : err}`);
      errorCount++;
    }
  }

  // Handle style swapping
  let styleSwapCount = 0;
  for (const style of styles) {
    try {
      // Use the scanned frame if available, otherwise fall back to selection
      const nodesToProcess = scannedFrame ? [scannedFrame] : figma.currentPage.selection;
      
      for (const node of nodesToProcess) {
        // Process any node type that might have styles
        const styleSwaps = await swapStylesInNode(node, style.name, sourceLibrary, targetLibrary);
        styleSwapCount += styleSwaps;
      }
    } catch (err) {
      errorDetails.push(`Error swapping style '${style.name}': ${err instanceof Error ? err.message : err}`);
      errorCount++;
    }
  }
  
  // Note: removeOverrides() is called during component swap above,
  // which completely removes all overrides and resets instances to library defaults
  console.log('✅ All component instances reset to target library defaults');

  // Check if nothing was swapped
  const totalSwapped = swapCount + styleSwapCount;
  if (totalSwapped === 0 && errorCount === 0) {
    figma.ui.postMessage({ type: 'swap-error', message: 'No matching component instances or styles found in selection.', details: [] });
    console.error('Swap error: No matching items found in selection.');
    return;
  }

  if (errorCount > 0) {
    figma.ui.postMessage({ type: 'swap-error', message: `Swap completed with ${errorCount} errors. ${totalSwapped} items swapped.`, details: errorDetails });
    console.error('Swap error details:', errorDetails);
  } else {
    figma.ui.postMessage({ type: 'swap-complete', message: `✅ Swap completed successfully! ${swapCount} components and ${styleSwapCount} styles swapped.` });
    figma.notify(`✅ Swap completed! ${swapCount} components and ${styleSwapCount} styles swapped.`);
  }
}

// ===== DELETE EVERYTHING BELOW THIS LINE! =====