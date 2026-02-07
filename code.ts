// Main Figma Swap Library Plugin (clean template)
console.log('🚀 Plugin starting... Version: Local-Sync-Only');
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
  id: string; // The ID of the Main Component (for swapping)
  instanceId?: string; // The ID of the Instance on canvas (for retrieving overrides)
  name: string;  
  displayName: string;  
  library: string;
  remote: boolean;
  parentName: string; 
  libraryFileId?: string; 
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
  // Ensure libraries are loaded before scanning
  await loadConnectedLibraries();
  
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
  variables?: Record<string, string>;
  thumbnail?: string;
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
          const varCount = lib.variables ? Object.keys(lib.variables).length : 0;
          console.log(`   📚 Library [${index}] "${lib.name}": ${compCount} components, ${styleCount} styles, ${varCount} variables.`);
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
  
  // We only support Local Sync now, so background refresh is limited.
  // We can only refresh the library that matches the CURRENT file.
  
  const currentFileName = figma.root.name;
  const matchingLibIndex = CONNECTED_LIBRARIES.findIndex(l => l.name === currentFileName && l.type === 'Local');
  
  if (matchingLibIndex >= 0) {
      console.log(`   - Auto-refreshing current file library: ${currentFileName}`);
      await handleSyncCurrentFile();
  }
}



function updateMappingsFromConnected() {
  // Reset mappings to defaults first to avoid stale data
  COMPONENT_KEY_MAPPING = { ...DEFAULT_COMPONENT_KEY_MAPPING };
  STYLE_KEY_MAPPING = { ...DEFAULT_STYLE_KEY_MAPPING };
  VARIABLE_KEY_MAPPING = { ...DEFAULT_VARIABLE_KEY_MAPPING };
  VARIABLE_ID_MAPPING = { ...DEFAULT_VARIABLE_ID_MAPPING };

  CONNECTED_LIBRARIES.forEach(lib => {
    if (lib.components) {
      COMPONENT_KEY_MAPPING[lib.name] = { ...COMPONENT_KEY_MAPPING[lib.name], ...lib.components };
    }
    if (lib.styles) {
      STYLE_KEY_MAPPING[lib.name] = { ...STYLE_KEY_MAPPING[lib.name], ...lib.styles };
    }
    if (lib.variables) {
      VARIABLE_KEY_MAPPING[lib.name] = { ...VARIABLE_KEY_MAPPING[lib.name], ...lib.variables };
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

// Check if current file is already synced
async function handleCheckCurrentFileStatus() {
  const name = figma.root.name;
  let fileId = figma.root.getPluginData('swap_library_file_id');
  
  if (!fileId) {
      if (figma.fileKey) {
          fileId = figma.fileKey;
      } else {
          // Check for existing by name (migration logic)
          const existingLib = CONNECTED_LIBRARIES.find(l => l.name === name && l.type === 'Local');
          if (existingLib) {
              fileId = existingLib.id;
          }
      }
  }

  const isSynced = fileId ? CONNECTED_LIBRARIES.some(l => l.id === fileId) : false;
  
  figma.ui.postMessage({
      type: 'CURRENT_FILE_STATUS',
      name: name,
      isSynced: isSynced
  });
}

// Sync Current File as Library
async function handleSyncCurrentFile() {
  console.log('🔄 Syncing current file as library...');
  // figma.notify("Syncing library...");
  
  try {
    const name = figma.root.name;
    
    // Use a persistent ID stored in Plugin Data if available, otherwise generate one
    let fileId = figma.root.getPluginData('swap_library_file_id');
    if (!fileId) {
        // If this is a published file, use the fileKey.
        if (figma.fileKey) {
            fileId = figma.fileKey;
        } else {
            // Migration: Check if we already have a local library with this name to avoid duplicates
            // This handles the case where a user syncs a file that was synced before we added persistent IDs
            const existingLib = CONNECTED_LIBRARIES.find(l => l.name === name && l.type === 'Local');
            if (existingLib) {
                fileId = existingLib.id;
                console.log(`   ℹ️ Found existing library by name: "${name}" -> Reusing ID: ${fileId}`);
            } else {
                fileId = `local_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
            }
        }
        figma.root.setPluginData('swap_library_file_id', fileId);
    }
    
    // Use the persistent ID as the library ID
    const fileKey = fileId;
    
    console.log(`   - File: ${name} (${fileKey})`);
    
    // 1. Scan Local Variables
    const variables: Record<string, string> = {};
    try {
        const localVars = await figma.variables.getLocalVariablesAsync();
        localVars.forEach(v => {
            variables[v.name] = v.key;
        });
        console.log(`   ✅ Found ${localVars.length} local variables.`);
    } catch (e) {
        console.error('   ❌ Failed to fetch local variables:', e);
    }

    // 2. Scan Local Styles
    const styles: Record<string, string> = {};
    try {
        // Use Async methods for styles to avoid "dynamic-page" errors
        const paintStyles = await figma.getLocalPaintStylesAsync();
        console.log(`   🔍 Debug: Found ${paintStyles.length} paint styles`);
        paintStyles.forEach(s => {
             // For Local Sync, we use ID for lookup, but we should also store the Key if possible for cross-file matching?
             // Actually, the issue seen in logs is that keys in storage have a trailing comma: "S:...,"
             // This is likely due to how we are storing them.
             // Wait, in the code above we are storing s.id.
             // If s.id contains a comma, that's weird.
             // Let's look at the storage dump again.
             // "styles": { "Primary": "S:82e4...," }
             // The ID of a local style usually looks like "S:key," or "S:key,local"
             // We should probably store the KEY if we want to match against other files, or ID if we want to match locally.
             // But the scan logic compares against node.fillStyleId (which is an ID) OR node.fillStyleId -> getStyleById -> style.key
             
             // If we store ID, we can match against node.fillStyleId directly.
             // If we store Key, we must resolve node style to key.
             
             // The scan logic does:
             // const paintStyle = await figma.getStyleByIdAsync(node.fillStyleId);
             // if (STYLE_KEY_MAPPING[lib][name] === paintStyle.key) ...
             
             // So the scan logic expects the mapping to contain KEYS.
             // But here in handleSyncCurrentFile, we are storing IDs: styles[s.name] = s.id;
             
             // And it seems s.id for a local style is "S:key," (with a comma).
             // But paintStyle.key (from getStyleById) is just "key" (no S:, no comma).
             
             // FIX: We should store s.key instead of s.id for the mapping, 
             // OR we need to clean up the ID if we really want to use IDs.
             // Since the scan logic compares against paintStyle.key, we MUST store s.key here.
             
             styles[s.name] = s.key; 
             // console.log(`     - Paint Style: ${s.name}, Key: ${s.key}`);
        });
        
        const textStyles = await figma.getLocalTextStylesAsync();
        textStyles.forEach(s => {
            styles[s.name] = s.key;
        });
        
        const effectStyles = await figma.getLocalEffectStylesAsync();
        effectStyles.forEach(s => {
             styles[s.name] = s.key;
        });
        
        const gridStyles = await figma.getLocalGridStylesAsync();
        gridStyles.forEach(s => {
             styles[s.name] = s.key;
        });
        
        console.log(`   ✅ Found ${Object.keys(styles).length} local styles.`);
    } catch (e) {
        console.error('   ❌ Failed to fetch local styles:', e);
    }

    // 3. Scan Local Components
    const components: Record<string, string> = {};
    try {
        // Must load all pages before using findAllWithCriteria in dynamic-page mode
        await figma.loadAllPagesAsync();
        
        const componentNodes = figma.root.findAllWithCriteria({ types: ['COMPONENT', 'COMPONENT_SET'] });
        console.log(`   🔍 Debug: Found ${componentNodes.length} component nodes`);
        
        componentNodes.forEach(node => {
            // Store KEY instead of ID for better cross-file matching
            // We will handle local lookup by scanning keys if needed
            let entryName = node.name;
            if (node.parent && node.parent.type === 'COMPONENT_SET') {
                entryName = `${node.parent.name}/${node.name}`;
            }
            components[entryName] = node.key;
            // console.log(`     - Component: ${entryName}, Key: ${node.key}`);
        });
        console.log(`   ✅ Found ${Object.keys(components).length} local components.`);
    } catch (e) {
        console.error('   ❌ Failed to fetch local components:', e);
    }

    // 3.5 Capture Thumbnail
    let thumbnailBase64: string | undefined;
    try {
        let thumbnailNode: SceneNode | null = null;
        
        // Strategy 0: Check for official file thumbnail (Best)
        try {
            const officialThumbnail = await figma.getFileThumbnailNodeAsync();
            if (officialThumbnail) {
                thumbnailNode = officialThumbnail;
                console.log('   🖼️ Using official file thumbnail node');
            }
        } catch (e) {
            console.log('   ℹ️ getFileThumbnailNodeAsync not supported or failed');
        }

        // Strategy 1: Check selection
        if (!thumbnailNode && figma.currentPage.selection.length === 1 && figma.currentPage.selection[0].type === 'FRAME') {
            thumbnailNode = figma.currentPage.selection[0];
            console.log('   🖼️ Using selected frame as thumbnail');
        }
        
        // Strategy 2: Check for "Cover" or "Thumbnail" page
        if (!thumbnailNode) {
            const coverPage = figma.root.children.find(p => 
                p.name.toLowerCase().includes('cover') || 
                p.name.toLowerCase().includes('thumbnail')
            );
            if (coverPage) {
                // Find first frame in cover page
                const frame = coverPage.children.find(n => n.type === 'FRAME');
                if (frame) {
                    thumbnailNode = frame;
                    console.log(`   🖼️ Found thumbnail frame in page "${coverPage.name}"`);
                }
            }
        }

        // Strategy 3: Use the first frame on the current page if nothing else found
        if (!thumbnailNode) {
             const firstFrame = figma.currentPage.children.find(n => n.type === 'FRAME');
             if (firstFrame) {
                 thumbnailNode = firstFrame;
                 console.log('   🖼️ Fallback: Using first frame on current page as thumbnail');
             }
        }

        if (thumbnailNode) {
            // Export to PNG
            const bytes = await (thumbnailNode as FrameNode).exportAsync({
                format: 'PNG',
                constraint: { type: 'SCALE', value: 0.5 } // Scale down to reduce size
            });
            
            // Convert to Base64
            // Use figma.base64Encode if available, otherwise manual fallback
            if (typeof figma.base64Encode === 'function') {
                const base64 = figma.base64Encode(bytes);
                thumbnailBase64 = `data:image/png;base64,${base64}`;
            } else {
                // Manual fallback using our helper function
                try {
                    const base64 = bufferToBase64(bytes);
                    thumbnailBase64 = `data:image/png;base64,${base64}`;
                } catch (e) {
                     console.warn('   ⚠️ Manual base64 encoding failed:', e);
                }
            }
            console.log('   ✅ Thumbnail captured and converted to Base64');
        }
    } catch (e) {
        console.warn('   ⚠️ Failed to capture thumbnail:', e);
    }

    // 4. Construct Library Object
    const newLib = {
        name: name,
        id: fileKey,
        key: fileKey,
        type: 'Local', // Mark as locally synced
        lastSynced: new Date().toISOString(),
        components: components,
        styles: styles,
        variables: variables,
        thumbnail: thumbnailBase64
    };

    // 5. Save to Storage
    // Remove existing if present (update)
    const existingIndex = CONNECTED_LIBRARIES.findIndex(l => l.id === fileKey);
    if (existingIndex >= 0) {
        CONNECTED_LIBRARIES[existingIndex] = newLib;
        // figma.notify(`Library "${name}" updated!`);
    } else {
        CONNECTED_LIBRARIES.push(newLib);
        // figma.notify(`Library "${name}" synced!`);
    }

    await saveConnectedLibraries();
    updateMappingsFromConnected();
    sendConnectedLibraries();

  } catch (err) {
    console.error('❌ Error syncing current file:', err);
    figma.notify('Failed to sync current file.');
  }
}






    

        

            


            // � RE-CHECK COLLECTIONS (The "Wake Up" Check)



            // �🕵️‍♂️ INSPECT IMPORTED STYLES FOR VARIABLES

            










        










async function handleRefreshLibrary(libraryId: string) {
  console.log(`🔄 Manual refresh requested for library ${libraryId}`);
  
  const libIndex = CONNECTED_LIBRARIES.findIndex(l => l.id === libraryId);
  if (libIndex === -1) {
      console.warn(`⚠️ Library ${libraryId} not found in connected libraries.`);
      return;
  }
  
  const lib = CONNECTED_LIBRARIES[libIndex];
  
  // Only refresh local libraries (since we removed remote support)
  if (lib.type === 'Local') {
      // For local libraries, we can't really "refresh" them from another file easily
      // unless we are IN that file.
      // If we are in the file that matches the library, we can re-sync it.
      if (figma.root.name === lib.name) {
          await handleSyncCurrentFile();
      } else {
          figma.notify(`To refresh "${lib.name}", please open that file and click "Sync Current File".`);
      }
  }
}

async function handleRemoveLibrary(libraryId: string) {
  CONNECTED_LIBRARIES = CONNECTED_LIBRARIES.filter(l => l.id !== libraryId);
  await saveConnectedLibraries();
  sendConnectedLibraries();
  figma.notify('Library removed.');
}



// Call load on start - Removed as it is now handled in the auto-scan IIFE
// loadConnectedLibraries();

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
    case 'RUN_DIAGNOSTICS':
      await runDiagnostics();
      break;
    case 'GET_TARGET_COLOR':
      await getTargetColor(msg.tokenId, msg.styleName, msg.sourceLibrary, msg.targetLibrary);
      break;
    case 'UPDATE_LIBRARY_DEFINITIONS':
      await syncLibraryDefinitions();
      break;
    case 'SYNC_CURRENT_FILE':
      await handleSyncCurrentFile();
      break;
    case 'CHECK_CURRENT_FILE_STATUS':
      await handleCheckCurrentFileStatus();
      break;
    case 'RESET_PLUGIN':
      await figma.clientStorage.setAsync('connected_libraries', []);
      CONNECTED_LIBRARIES = [];
      sendConnectedLibraries();
      figma.notify('Plugin data reset');
      await handleScanFrames();
      break;
    case 'REMOVE_LIBRARY':
      await handleRemoveLibrary(msg.libraryId);
      break;
    case 'REFRESH_LIBRARY':
      await handleRefreshLibrary(msg.id);
      break;
    case 'GET_CONNECTED_LIBRARIES':
      console.log('📩 Received GET_CONNECTED_LIBRARIES request');
      sendConnectedLibraries();
      break;
    case 'GET_COMPONENT_PROPERTIES':
      await handleGetComponentProperties(msg.sourceId, msg.targetKey, msg.targetName, msg.targetLibraryName, msg.sourceLibraryName);
      break;
    case 'SHOW_NATIVE_TOAST':
      figma.notify(msg.message || 'Swap completed successfully!');
      break;
    case 'close-plugin':
      figma.closePlugin();
      break;
    case 'DEBUG_DUMP_STORAGE':
      console.log('🔍 DEBUG: Dumping Client Storage...');
      try {
        const stored = await figma.clientStorage.getAsync('connected_libraries');
        console.log('📦 FULL STORAGE DUMP:', JSON.stringify(stored, null, 2));
        figma.notify('Storage dumped to console.');
      } catch (e) {
        console.error('❌ Failed to dump storage:', e);
      }
      break;
    case 'CLEAR_LIBRARIES':
      CONNECTED_LIBRARIES = [];
      await saveConnectedLibraries();
      updateMappingsFromConnected();
      sendConnectedLibraries();
      figma.notify('All libraries cleared.');
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
          // Check if library is Local
          const libObj = CONNECTED_LIBRARIES.find(l => l.name === libName);
          const isLocal = libObj?.type === 'Local';

          for (const compName of Object.keys(COMPONENT_KEY_MAPPING[libName])) {
            const mappedKey = COMPONENT_KEY_MAPPING[libName][compName];
            
            // Debug logging for key matching
            // console.log(`🔍 Checking ${compName} in ${libName}: mapped=${mappedKey}, actual=${component.key}, id=${componentId}`);
            
            // Match scoring:
            // 5. Exact key match (or ID match for Local)
            // 4. Component ID match (last part of key)
            // 3. Mapped key is contained in component key (for some remote library formats)
            // 2. Component key is contained in mapped key (reverse check)
            // 1. Name match (fallback) - CAUTION: This can be risky if names are not unique
            
            let score = 0;
            
            if (isLocal) {
                // For Local libraries, mappedKey is the ID. Compare with component.id
                // Also check component.key because we recently switched to storing keys for local libs
                if (mappedKey === component.id) score = 5;
                else if (mappedKey === component.key) score = 5;
                else if (component.name === compName) score = 1;
            } else {
                // For Remote libraries, mappedKey is the Key.
                if (mappedKey === component.key) score = 5;
                else if (mappedKey === componentId) score = 4;
                else if (component.key.includes(mappedKey)) score = 3;
                else if (mappedKey.includes(componentId)) score = 2;
                else if (component.name === compName) score = 1;
            }

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
              instanceId: node.type === 'INSTANCE' ? node.id : undefined,
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
  // Skip style scanning if the node is part of an instance (unless it's an override?)
  // Actually, the requirement is "scan list should not show any styles that are applied to a component".
  // This likely means: if I have an instance, don't scan its children for styles.
  // But if I have a frame with a rectangle in it, scan the rectangle.
  // If I have an instance, I might want to scan the instance itself (e.g. if the instance has a fill style override on the top level).
  // But I should NOT scan the children of the instance.
  
  // scanNodeForAssets calls scanNodeForStyles(node) AND then recurses into children.
  // If node is an INSTANCE, we should scan it for styles (in case the instance itself has a style),
  // BUT we should NOT recurse into its children if we want to exclude "styles applied to a component" (meaning internal styles).
  
  // Only scan for styles if it's NOT an instance (User requirement: ignore styles applied to component instances)
  if (node.type !== 'INSTANCE') {
      await scanNodeForStyles(node, tokens);
  }
  
  // Only recurse if NOT an instance
  if (node.type !== 'INSTANCE' && 'children' in node && node.children) {
    for (const child of node.children) {
      await scanNodeForAssets(child, components, tokens);
    }
  }
}

// Scan a node for style tokens
async function scanNodeForStyles(node: SceneNode, tokens: TokenInfo[]): Promise<void> {
  // console.log(`🔍 Scanning node: ${node.name} (type: ${node.type})`);
  
  // Check for paint styles (traditional Shark approach)
  // Note: INSTANCE nodes do not have 'fillStyleId' directly on them in the Figma API types,
  // but they can have fill overrides. However, accessing 'fillStyleId' on an InstanceNode
  // might not be straightforward if it's not in the type definition.
  // Actually, InstanceNode DOES have fillStyleId if it has a fill override.
  // But we need to be careful.
  
  if ('fillStyleId' in node && node.fillStyleId && typeof node.fillStyleId === 'string' && (node.fillStyleId as string).length > 0) {
    try {
      const paintStyle = await figma.getStyleByIdAsync(node.fillStyleId as string);
      if (paintStyle && paintStyle.type === 'PAINT') {
        // console.log(`🎨 Found paint style: ${paintStyle.name}, key: ${paintStyle.key}`);
        // Determine library from style key
        let library = 'Unknown';
        for (const libName of Object.keys(STYLE_KEY_MAPPING)) {
          for (const styleName of Object.keys(STYLE_KEY_MAPPING[libName])) {
            const mappedKey = STYLE_KEY_MAPPING[libName][styleName];
            const currentKey = paintStyle.key;
            
            // Normalize keys for comparison (strip S: prefix if present, and trailing comma)
            const normMapped = mappedKey.replace(/^S:/, '').replace(/,$/, '');
            const normCurrent = currentKey.replace(/^S:/, '').replace(/,$/, '');
            
            // console.log(`Checking style ${styleName} in ${libName}: ${mappedKey} vs ${currentKey} (Norm: ${normMapped} vs ${normCurrent})`);
            
            if (mappedKey === currentKey || normMapped === normCurrent) {
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
            const mappedKey = STYLE_KEY_MAPPING[libName][styleName];
            const currentKey = textStyle.key;
            
            // Normalize keys for comparison (strip S: prefix if present, and trailing comma)
            const normMapped = mappedKey.replace(/^S:/, '').replace(/,$/, '');
            const normCurrent = currentKey.replace(/^S:/, '').replace(/,$/, '');

            if (mappedKey === currentKey || normMapped === normCurrent) {
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
            const mappedKey = STYLE_KEY_MAPPING[libName][styleName];
            const currentKey = effectStyle.key;
            
            // Normalize keys for comparison (strip S: prefix if present, and trailing comma)
            const normMapped = mappedKey.replace(/^S:/, '').replace(/,$/, '');
            const normCurrent = currentKey.replace(/^S:/, '').replace(/,$/, '');

            if (mappedKey === currentKey || normMapped === normCurrent) {
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

  // Check if target library uses styles
  // Priority: 1. Dynamic Metadata 2. Global Mapping
  const targetLibObj = CONNECTED_LIBRARIES.find(l => l.name === normalizedTargetLibrary);
  let targetStyleKey: string | undefined;
  
  if (targetLibObj && targetLibObj.styles) {
      targetStyleKey = targetLibObj.styles[styleName];
  }
  
  if (!targetStyleKey) {
      targetStyleKey = STYLE_KEY_MAPPING[normalizedTargetLibrary]?.[styleName];
  }

  console.log(`📍 Target library: ${normalizedTargetLibrary}, Style key: ${targetStyleKey}`);

  if (targetStyleKey) {
      try {
          const style = await figma.importStyleByKeyAsync(targetStyleKey);
          if (style && style.type === 'PAINT') {
              const paintStyle = style as PaintStyle;
              if (paintStyle.paints.length > 0) {
                  const color = getColorValue(paintStyle.paints[0]);
                  console.log(`🎨 Style color resolved to: ${color}`);
                  figma.ui.postMessage({ type: 'TARGET_COLOR_RESULT', tokenId, color });
                  return;
              }
          }
      } catch (e) {
          console.warn(`❌ Failed to import style ${styleName}:`, e);
      }
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
        
        if (expectedKey) {
            let match = false;
            
            // Check Key (Standard)
            if (mainComponent.key === expectedKey) {
                match = true;
            }
            // Check ID (Legacy/Local fallback)
            else if (mainComponent.id === expectedKey) {
                match = true;
            }
            // Check Name (Fallback for detached/renamed/mismatched keys)
            else if (mainComponent.name === name) {
                 console.log(`⚠️ Key mismatch for ${name}. Expected: ${expectedKey}, Found: ${mainComponent.key}. Matching by name.`);
                 match = true;
            }

            if (match) {
                console.log(`✅ Match found! Adding instance.`);
                found.push(node);
            }
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

// Helper to import style by key OR get local style by ID
async function importOrGetStyle(keyOrId: string, libraryName: string): Promise<BaseStyle | null> {
    const lib = CONNECTED_LIBRARIES.find(l => l.name === libraryName);
    if (lib && lib.type === 'Local') {
        try {
            // For Local libraries, we stored the ID in the mapping
            const style = figma.getStyleById(keyOrId);
            if (style) return style;
        } catch (e) {
            // Ignore error, try import
        }
    }
    try {
        return await figma.importStyleByKeyAsync(keyOrId);
    } catch (e) {
        console.warn(`Failed to import style ${keyOrId}: ${e}`);
        return null;
    }
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
  // Priority: 1. Dynamic Metadata (targetLibObj.variables) 2. Global Mapping
  let targetVariableKey = targetLibObj?.variables?.[styleName];
  if (!targetVariableKey) {
      targetVariableKey = VARIABLE_KEY_MAPPING[targetLibrary]?.[styleName];
  }
  
  const targetVariableId = VARIABLE_ID_MAPPING[targetLibrary]?.[styleName];

  console.log(`  📋 Processing node: ${node.name} (type: ${node.type}) looking for style: ${styleName}`);
  console.log(`  ℹ️ Target Style Key for '${styleName}': ${targetStyleKey}`);
  console.log(`  ℹ️ Target Variable Key for '${styleName}': ${targetVariableKey}`);

  // Process fill styles on any node that has them (including FRAME)
  if ('fillStyleId' in node && node.fillStyleId && typeof node.fillStyleId === 'string') {
    try {
      const currentStyle = await figma.getStyleByIdAsync(node.fillStyleId);
      if (currentStyle && currentStyle.key === sourceStyleKey) {
        console.log(`  ✅ Found ${styleName} style on ${node.type}`);
        
        // Check for target style first (if target is a style library)
        // Use the targetStyleKey we already looked up from metadata or mapping
        if (targetStyleKey) {
            try {
                const targetStyle = await importOrGetStyle(targetStyleKey, targetLibrary);
                if (targetStyle) {
                    await node.setFillStyleIdAsync(targetStyle.id);
                    console.log(`  ✅ Swapped fill style to '${styleName}' on ${node.type}`);
                    swapCount++;
                    return swapCount; // Done swapping this fill
                }
            } catch (e) {
                console.warn(`  ❌ Failed to import fill style ${styleName}:`, e);
            }
        }
        
        // If we have a target variable KEY or ID, bind to it (Legacy/Variable logic)
        // OR if we didn't find a style, try to find a variable by name
        // OR if we are explicitly trying to swap styles to variables (Shark -> Monkey scenario)
        if (targetVariableKey || targetVariableId || !targetStyleKey) {
          try {
            console.log(`  🔄 Checking for variable in ${targetLibrary}: ${styleName}`);
            const nodeAny = node as any;
            
            // Try setting the fill with variable binding directly
            if (Array.isArray(nodeAny.fills) && nodeAny.fills.length > 0) {
              try {
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
                    // CRITICAL FIX: We must search ALL available variables, not just local ones.
                    // getLocalVariablesAsync() only returns variables DEFINED in this file or explicitly imported.
                    // If a library is enabled but its variables haven't been used yet, they won't be in getLocalVariablesAsync().
                    // We need to use importVariableByKeyAsync if we have the key (which we should from metadata).
                    
                    // If we don't have a key, we are in trouble. But wait, we should have keys from the library metadata.
                    // If targetVariableKey is missing, it means the library metadata is incomplete (0 variables found).
                    
                    // If the library is enabled, we can try to find the variable in the available collections?
                    // No, there is no API to "search all available variables" without importing them.
                    
                    // However, if we are here, it means we failed to find the variable by key.
                    // This usually happens if the library is NOT enabled or the key is wrong.
                    
                    const localVars = await figma.variables.getLocalVariablesAsync();
                    // Filter variables to try and match the target library if possible, or just match name
                    // We prioritize variables that might be from the target library (checking collection name)
                    
                    const candidates = localVars.filter(v => v.name === styleName);
                    
                    if (candidates.length > 0) {
                        // Try to find one from a collection that matches target library name
                        for (const v of candidates) {
                            try {
                                const collection = await figma.variables.getVariableCollectionByIdAsync(v.variableCollectionId);
                                if (collection && collection.name.includes(targetLibrary)) {
                                    variable = v;
                                    console.log(`  ✅ Found variable "${v.name}" in collection "${collection.name}" (matches target library)`);
                                    break;
                                }
                            } catch (e) {
                                // ignore
                            }
                        }
                        
                        // If no specific collection match, just take the first one (or maybe prefer remote?)
                        if (!variable) {
                             variable = candidates[0];
                             console.log(`  ⚠️ Found variable "${variable.name}" but collection didn't match "${targetLibrary}". Using it anyway.`);
                        }
                    } else {
                        // Last Resort: Try to find a variable that *contains* the style name (e.g. "Brand/Primary" matches "Primary")
                        // This helps when mapping flat styles to nested variables
                        const looseCandidates = localVars.filter(v => v.name.endsWith(`/${styleName}`) || v.name === styleName);
                        if (looseCandidates.length > 0) {
                             variable = looseCandidates[0];
                             console.log(`  ✅ Found variable by loose name match: "${variable.name}"`);
                        }
                    }
                    
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
                  
                  // FINAL FALLBACK: If we can't find the variable, maybe the user hasn't enabled the library yet?
                  // We should notify them.
                  if (targetVariableKey === undefined) {
                       console.warn(`  ⚠️ Hint: Variable '${styleName}' key is undefined. Is the target library enabled in Assets?`);
                  }
                }
              } catch (bindErr) {
                console.log(`  ❌ Error binding fill variable: ${bindErr}`);
              }
            }
          } catch (e) {
            console.log(`  ℹ️ Could not bind variable directly: ${e}`);
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
                    const targetStyle = await importOrGetStyle(targetStyleKey, targetLibrary);
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
        
        // Check for target style first (if target is a style library)
        const targetStyleKey = STYLE_KEY_MAPPING[normalizedTargetLibrary]?.[styleName];
        if (targetStyleKey) {
            try {
                const targetStyle = await importOrGetStyle(targetStyleKey, targetLibrary);
                if (targetStyle) {
                    await node.setStrokeStyleIdAsync(targetStyle.id);
                    console.log(`  ✅ Swapped stroke style to '${styleName}' on ${node.type}`);
                    swapCount++;
                    return swapCount; // Done swapping this stroke
                }
            } catch (e) {
                console.warn(`  ❌ Failed to import stroke style ${styleName}:`, e);
            }
        }

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
            const targetStyle = await importOrGetStyle(targetStyleKey, targetLibrary);
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
        } else {
            console.warn(`  ❌ No target style key found for ${styleName}`);
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
async function restoreTargetStyles(instanceNode: SceneNode, targetComponent: ComponentNode, targetLibrary: string): Promise<void> {
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
    async function applyTargetStyles(node: SceneNode): Promise<void> {
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
            // const allLibraryFiles = figma.getSharedLibraryFiles();
            
            // for (const libFile of allLibraryFiles) {
            //   console.log(`📚 Checking library file: ${libFile.name}`);
              
            //   // Get all paint styles from this library
            //   try {
            //     const paintStyles = libFile.getSharedPluginData('figma', 'paintStyles');
            //     if (paintStyles) {
            //       const styles = JSON.parse(paintStyles);
            //       console.log(`  Found styles:`, styles);
            //     }
            //   } catch (e) {
            //     console.warn(`  Could not get styles from library:`, e);
            //   }
            // }
            
            // Also try getLocalPaintStylesAsync to search for the style
            const allPaintStyles = figma.getLocalPaintStyles();
            console.log(`🔎 Searching in ${allPaintStyles.length} local styles for "${styleName}"`);
            
            for (const style of allPaintStyles) {
              console.log(`  Style: ${style.name} (id: ${style.id})`);
              if (style.name === styleName) {
                console.log(`✅ Found matching style! Applying to "${node.name}"`);
                await inst.setFillStyleIdAsync(style.id);
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
          await applyTargetStyles(child);
        }
      }
    }
    
    // Apply styles to instance
    await applyTargetStyles(instanceNode);
  } catch (error) {
    console.warn('Error applying target styles:', error);
  }
}

// Add detailed swap logic for PERFORM_LIBRARY_SWAP
async function performLibrarySwap(components: any[], styles: any[], sourceLibrary: string, targetLibrary: string) {
  console.log('🔄 performLibrarySwap called!');
  console.log(`Components to swap: ${components.length}`);
  console.log(`Styles to swap: ${styles.length}`);
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
      console.log(`🔄 Processing component: ${comp.name} -> Target: ${comp.targetName} (Key: ${comp.targetKey})`);
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
            let targetKey: string | undefined = comp.targetKey;
            
            if (!targetKey && targetLibObj && targetLibObj.components) {
                // Use explicitly provided target name if available, otherwise fallback to source name
                const lookupName = comp.targetName || comp.name;
                
                // console.log(`  🔎 Lookup: '${lookupName}' in library '${targetLibrary}'`);
                
                targetKey = targetLibObj.components[lookupName];
                
                if (targetKey) {
                    console.log(`  ✅ Found exact match: ${lookupName} -> ${targetKey}`);
                }
                
                // Fallback: If not found, try looking up the parent Component Set
                // This handles cases where we are swapping a specific variant but the target library only indexed the Component Set,
                // or if the exact variant mapping is missing.
                if (!targetKey) {
                   // 1. Try using source parent name if available (assuming same structure in target)
                   if (comp.parentName && targetLibObj.components[comp.parentName]) {
                       targetKey = targetLibObj.components[comp.parentName];
                       console.log(`  ✅ Fallback: Found parent Component Set directly: ${comp.parentName} -> ${targetKey}`);
                   }
                   // 2. Try splitting lookupName by '/' (for namespaced variants like "Button/Type=Primary")
                   else if (lookupName.includes('/')) {
                       const parentPart = lookupName.split('/')[0];
                       if (targetLibObj.components[parentPart]) {
                           targetKey = targetLibObj.components[parentPart];
                           console.log(`  ✅ Fallback: Found parent Component Set by path: ${parentPart} -> ${targetKey}`);
                       }
                   }
                }

                if (!targetKey) {
                    console.warn(`  ⚠️ Component '${lookupName}' not found in target library '${targetLibrary}'.`);
                }
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
              let importedComponent: ComponentNode | ComponentSetNode | null = null;
              
              // Check if target is the Current File (Local Context)
              // We check both name and File ID/Key if available
              const isTargetCurrentFile = figma.root.name === targetLibrary || 
                                          (targetLibObj?.id && figma.root.getPluginData('swap_library_file_id') === targetLibObj.id);
              
              if (isTargetCurrentFile) {
                  // Try to find by ID first (Legacy support)
                  const localNode = figma.getNodeById(targetKey);
                  if (localNode && (localNode.type === 'COMPONENT' || localNode.type === 'COMPONENT_SET')) {
                      importedComponent = localNode as ComponentNode | ComponentSetNode;
                      console.log(`  ✅ Found local component by ID: ${importedComponent.name}`);
                  } else {
                      // Try to find by Key (New support)
                      // This is expensive, so we should optimize if possible, but for now scan all components
                      // findAllWithCriteria is faster than findAll
                      const allComponents = figma.root.findAllWithCriteria({ types: ['COMPONENT', 'COMPONENT_SET'] });
                      const found = allComponents.find(c => c.key === targetKey);
                      if (found) {
                          importedComponent = found;
                          console.log(`  ✅ Found local component by Key: ${found.name}`);
                      }
                  }
              } else {
                  // Target is Remote
                  console.log(`  🌐 Attempting to import REMOTE component with key: ${targetKey}`);
                  try {
                      importedComponent = await figma.importComponentByKeyAsync(targetKey);
                      console.log(`  ✅ Successfully imported remote component: ${importedComponent.name}`);
                  } catch (e) {
                      console.warn(`  ❌ Failed to import component by key ${targetKey}:`, e);
                      
                      // NEW: Retry with the specific Variant Name (ignoring targetName override)
                      // If targetName set the lookup to the Component Set (e.g. "Rectangle") and it failed,
                      // we should try the specific variant name from the source (e.g. "Rectangle/Shape=Square")
                      // provided that the target library has a key for it.
                      if (!importedComponent && comp.name && comp.name !== comp.targetName) {
                          if (targetLibObj && targetLibObj.components && targetLibObj.components[comp.name]) {
                               const variantKey = targetLibObj.components[comp.name];
                               if (variantKey && variantKey !== targetKey) {
                                   console.log(`  🔄 Retry: Attempting to import by Source Variant Name: ${comp.name} (${variantKey})`);
                                   try {
                                       importedComponent = await figma.importComponentByKeyAsync(variantKey);
                                       if (importedComponent) {
                                           console.log(`  ✅ Successfully imported by Source Variant Name: ${importedComponent.name}`);
                                       }
                                   } catch (eV) {
                                       console.warn(`  ❌ Failed to import by Source Variant Name key:`, eV);
                                   }
                               }
                          }
                      }
                      
                      // Retry with Parent Component Set Key if this was a variant specific key
                      // Log shows expected lookup was "Rectangle/Shape=Square", but maybe that key refers to an unpublished/private main component?
                      // If we can default to the "Rectangle" set key, we might succeed.
                      
                      // Derive parent name if needed (comp.parentName might be undefined)
                      const parentName = comp.parentName || (comp.name.includes('/') ? comp.name.split('/')[0] : null);
                      
                      if (!importedComponent && parentName && targetLibObj && targetLibObj.components && targetLibObj.components[parentName]) {
                           const parentKey = targetLibObj.components[parentName];
                           if (parentKey && parentKey !== targetKey) {
                               console.log(`  🔄 Retry: Attempting to import Parent Component Set instead: ${parentName} (${parentKey})`);
                               try {
                                   importedComponent = await figma.importComponentByKeyAsync(parentKey);
                                   console.log(`  ✅ Successfully imported Parent Component Set: ${importedComponent.name}`);
                               } catch (e2) {
                                   console.warn(`  ❌ Failed to import Parent Component Set by key ${parentKey}:`, e2);
                               }
                           }
                      }
                      
                      // If still failing, try simple name splits (e.g. "Rectangle" from "Rectangle/Shape=Square")
                      if (!importedComponent && comp.name.includes('/')) {
                          const simpleName = comp.name.split('/')[0];
                          if (targetLibObj && targetLibObj.components && targetLibObj.components[simpleName]) {
                               const simpleKey = targetLibObj.components[simpleName];
                               if (simpleKey && simpleKey !== targetKey) {
                                   console.log(`  🔄 Retry: Attempting to import by Simple Name: ${simpleName} (${simpleKey})`);
                                   try {
                                       importedComponent = await figma.importComponentByKeyAsync(simpleKey);
                                       console.log(`  ✅ Successfully imported by Simple Name: ${importedComponent.name}`);
                                   } catch (e3) {
                                       console.warn(`  ❌ Failed to import by Simple Name key:`, e3);
                                   }
                               }
                          }
                      }
                  }
              }

              if (!importedComponent) {
                  throw new Error(`Could not find/import component: ${comp.name}`);
              }

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
                      // console.log(`  ⚙️ Captured overrides for nested instance #${nestedInstanceIndex}`);
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
                  if (instNode.counterAxisAlignItems !== 'MIN') layoutProps.counterAxisAlignItems = instNode.counterAxisAlignItems;
                  if (instNode.counterAxisSizingMode !== 'AUTO') layoutProps.counterAxisSizingMode = instNode.counterAxisSizingMode;
                  // Always capture width and height for resizing
                  layoutProps.width = instNode.width;
                  layoutProps.height = instNode.height;
                  if (Object.keys(layoutProps).length > 0) {
                    nestedInstanceLayouts.set(nestedInstanceIndex, layoutProps);
                    // console.log(`  📐 Captured layout for nested instance #${nestedInstanceIndex}`);
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
                  // console.log(`  📝 Captured text #${textNodeIndex - 1} from "${nodePath}"`);
                }
                if ('children' in node) {
                  for (const child of node.children) {
                    await captureTextValues(child, nodePath);
                  }
                }
              }
              // console.log('🔄 Capturing text values before swap...');
              await captureTextValues(instanceNode);
              // console.log(`  Total text nodes captured: ${textValues.size}`);
              
              // Handle Component Sets (swapComponent requires a ComponentNode)
              let componentToSwap: ComponentNode | null = null;
              
              if (importedComponent.type === 'COMPONENT_SET') {
                   console.log(`  ℹ️ Imported object is a Component Set: ${importedComponent.name}`);
                   // Try to find the matching variant by name parsing
                   // Source Name: "Rectangle/Shape=Square" -> Variant Property: "Shape=Square"
                   // Or simply "Shape=Square"
                   
                   // Parse desired properties from lookup name
                   // If lookupName is "Button/Type=Primary, State=Hover"
                   // We want to match { Type: "Primary", State: "Hover" }
                   
                   const targetVariantName = (comp.targetName || comp.name).split('/').pop(); // "Shape=Square"
                   
                   // Try to find a child with matching name
                   // Component variants have names like "Type=Primary, State=Hover"
                   
                   // 1. Try exact name match on children
                   const exactMatch = importedComponent.children.find(c => c.name === targetVariantName) as ComponentNode;
                   if (exactMatch) {
                       componentToSwap = exactMatch;
                       console.log(`  ✅ Found variant by exact name match in set: ${componentToSwap.name}`);
                   }
                   
                   // 2. Try parsing properties
                   if (!componentToSwap && targetVariantName) {
                       // Convert "Shape=Square, Size=Small" to object { Shape: "Square", Size: "Small" }
                       const desiredProps: Record<string, string> = {};
                       targetVariantName.split(',').forEach((pair: string) => {
                           const [k, v] = pair.split('=').map((s: string) => s.trim());
                           if (k && v) desiredProps[k] = v;
                       });
                       
                       if (Object.keys(desiredProps).length > 0) {
                           // Find variant that matches all desired properties
                           
                           const bestVariant = importedComponent.children.find(child => {
                               if (child.type !== 'COMPONENT') return false;
                               // Check if all desired props match this child's name properties
                               // Child name format: "Key1=Value1, Key2=Value2"
                               const childProps: Record<string, string> = {};
                               child.name.split(',').forEach((pair: string) => {
                                   const [k, v] = pair.split('=').map((s: string) => s.trim());
                                   if (k && v) childProps[k] = v;
                               });
                               
                               // Check match
                               return Object.entries(desiredProps).every(([k, v]) => childProps[k] === v);
                           });
                           
                           if (bestVariant) {
                               componentToSwap = bestVariant as ComponentNode;
                               console.log(`  ✅ Found variant by property match: ${componentToSwap.name}`);
                           }
                       }
                   }
                   
                   // 3. Fallback to default variant
                   if (!componentToSwap) {
                       if (importedComponent.defaultVariant) {
                           componentToSwap = importedComponent.defaultVariant;
                           console.log(`  ⚠️ Exact variant not found, using default variant: ${componentToSwap.name}`);
                       } else if (importedComponent.children.length > 0) {
                           componentToSwap = importedComponent.children[0] as ComponentNode;
                           console.log(`  ⚠️ Exact variant not found, using first available variant: ${componentToSwap.name}`);
                       }
                   }
              } else {
                  // It's already a specific component
                  componentToSwap = importedComponent as ComponentNode;
              }
              
              if (!componentToSwap) {
                  throw new Error(`Content imported but could not resolve to a specific component/variant.`);
              }

              console.log(`  ✅ Ready to swap instance to: ${componentToSwap.name}`);
              
              // Capture Mapped Properties (Before Swap)
              const mappedValues: Record<string, any> = {};
              if (comp.propertyMapping && Object.keys(comp.propertyMapping).length > 0) {
                  console.log(`  🔄 [Backend] Processing property mapping for ${comp.name}:`, comp.propertyMapping);
                  const currentProps = instanceNode.componentProperties;
                  // console.log('    [Backend] Current Instance Props:', JSON.stringify(currentProps));

                  for (const [targetProp, sourceProp] of Object.entries(comp.propertyMapping)) {
                      const sourceKeyRaw = sourceProp as string;
                      let valueToTransfer = undefined;
                      let matchedSourceKey = '';
                      
                      // 1. Direct key match (e.g. "State#123")
                      if (currentProps[sourceKeyRaw]) {
                          valueToTransfer = currentProps[sourceKeyRaw].value;
                          matchedSourceKey = sourceKeyRaw;
                      } 
                      // 2. Name match (e.g. sourceProp="State#123", key="State#456")
                      else if (sourceKeyRaw.includes('#')) {
                          const cleanSource = sourceKeyRaw.split('#')[0];
                          const match = Object.entries(currentProps).find(([k,v]) => k.split('#')[0] === cleanSource);
                          if (match) {
                              valueToTransfer = match[1].value;
                              matchedSourceKey = match[0];
                          }
                      } else {
                          // 3. Simple name match (sourceProp="State")
                          const match = Object.entries(currentProps).find(([k,v]) => k.split('#')[0] === sourceKeyRaw);
                          if (match) {
                              valueToTransfer = match[1].value;
                              matchedSourceKey = match[0];
                          }
                      }
                      
                      if (valueToTransfer !== undefined) {
                          mappedValues[targetProp] = valueToTransfer;
                          console.log(`    ✅ [Backend] Mapped '${sourceProp}' (found as ${matchedSourceKey}) -> '${targetProp}' = ${valueToTransfer}`);
                      } else {
                          console.warn(`    ⚠️ [Backend] Could not find source property '${sourceProp}' on instance.`);
                      }
                  }
              }

              // Perform the swap
              instance.swapComponent(componentToSwap);
              console.log(`✅ Swapped component: ${instance.name}`);

              // Remove all overrides so instance uses only target library defaults
              // Only do this if "Preserve style overrides" is NOT checked
              // IMPORTANT: Do this BEFORE applying mapped properties, otherwise we wipe them out!
              if (!comp.preserveStyle) {
                  try {
                    const inst = instanceNode as any;
                    if (typeof inst.removeOverrides === 'function') {
                      inst.removeOverrides();
                      console.log('✅ Removed all overrides (User requested reset)');
                    } else if (typeof inst.resetAllOverrides === 'function') {
                      // Fallback to resetAllOverrides if removeOverrides doesn't exist
                      inst.resetAllOverrides();
                      console.log('✅ Reset all overrides (fallback)');
                    }
                  } catch (e) {
                    console.warn('Error removing overrides:', e);
                  }
              } else {
                  console.log('ℹ️ Preserving style overrides (User requested)');
              }

              // Apply Mapped Properties (After Swap & Reset)
              let propertiesApplied = false;
              // Hoist finalPropsMap so it can be used in reapplyTextValues
              const finalPropsMap: Record<string, any> = {};

              // [FIX] Enforce specific variant properties
              // When swapping between variants, Figma might try to preserve the OLD variant properties (e.g. Shape=Round)
              // even if we swapped to a specific new variant (Shape=Square).
              // We must explicitly enforce the properties of the target variant on the new instance.
              if (componentToSwap.type === 'COMPONENT' && componentToSwap.parent && componentToSwap.parent.type === 'COMPONENT_SET') {
                  const variantProps = componentToSwap.variantProperties;
                  if (variantProps) {
                      console.log('  🔒 Enforcing variant properties:', JSON.stringify(variantProps));
                      // Match variant properties to instance property IDs
                      const instanceProps = instanceNode.componentProperties;
                      Object.entries(variantProps).forEach(([vName, vValue]) => {
                          const match = Object.entries(instanceProps).find(([k, v]) => k.split('#')[0] === vName);
                          if (match) {
                              const [propId, currentVal] = match;
                              // Only add if it's different or we want to be sure
                              finalPropsMap[propId] = vValue;
                          }
                      });
                  }
              }

              if (Object.keys(mappedValues).length > 0) {
                  console.log('  🔄 Applying mapped properties...');
                  const newProps = instanceNode.componentProperties;
                  // console.log('    [Backend] Debug: New Instance Props available:', Object.keys(newProps));
                  
                  for (const [targetKeyRaw, value] of Object.entries(mappedValues)) {
                       let finalKey = targetKeyRaw;
                       
                       // Verify key exists on new instance, or find name equivalent
                       if (!newProps[finalKey]) {
                            const cleanTarget = targetKeyRaw.split('#')[0];
                            console.log(`    ⚠️ Key '${targetKeyRaw}' not found on new instance. Searching for '${cleanTarget}'...`);
                            
                            const match = Object.entries(newProps).find(([k,v]) => k.split('#')[0] === cleanTarget);
                            if (match) {
                                finalKey = match[0];
                                console.log(`      ✅ Found match by name: '${finalKey}'`);
                            }
                            else {
                                console.warn(`      ❌ No match found for property '${cleanTarget}' on new instance.`);
                                finalKey = '';
                            }
                       } else {
                           // console.log(`    ✅ Exact key match: ${finalKey}`);
                       }
                       
                       if (finalKey) {
                           finalPropsMap[finalKey] = value;
                       }
                  }
              }
                  
              if (Object.keys(finalPropsMap).length > 0) {
                      try {
                          instanceNode.setProperties(finalPropsMap);
                          console.log('  ✅ Applied mapped properties success. Keys:', Object.keys(finalPropsMap));
                          console.log('  🔍 Verifying values stuck...');
                          const checkProps = instanceNode.componentProperties;
                          Object.keys(finalPropsMap).forEach(k => {
                              // console.log(`    - ${k.split('#')[0]}: Expected "${finalPropsMap[k]}", Got "${checkProps[k]?.value}"`);
                              if (checkProps[k]?.value != finalPropsMap[k]) {
                                  console.warn(`    ⚠️ Mismatch! Property ${k} did not update.`);
                              }
                          });
                          propertiesApplied = true;

                          // [DEBUG FIX] For Text Properties, explicitly finding nodes that use this property 
                          // and forcing their characters to update if they didn't.
                          // Figma sometimes needs a nudge if the property is 'consumed' but the text node doesn't redraw.
                          // However, setProperties SHOULD handle this. 
                          // Let's verify if the text node characters actuaally match the property value.
                          
                          // We can't easily find which text node is bound to which property without iterating.
                          // But we can iterate text nodes and check references.
                          instanceNode.findAll(n => n.type === 'TEXT').forEach(n => {
                             const textNode = n as TextNode;
                             if (textNode.componentPropertyReferences?.characters) {
                                 const propId = textNode.componentPropertyReferences.characters;
                                 if (finalPropsMap[propId] !== undefined) {
                                     const expected = finalPropsMap[propId];
                                     // This is purely for debug logging to confirm state
                                     if (textNode.characters !== String(expected)) {
                                         console.warn(`    ⚠️ Text Node "${textNode.name}" characters ("${textNode.characters}") do not match mapped property value ("${expected}"). Attempting force update...`);
                                     } 
                                 }
                             }
                          });

                      } catch (e) {
                          console.warn('  ⚠️ Failed to apply mapped properties:', e);
                      }
              }
              
              // Always attempt to restore content (Text) and Layout, but be smart about it.
              // Note: Style overrides are handled by removeOverrides() above based on comp.preserveStyle.
              
              // Reapply captured text values using same index-based approach
              console.log('🔄 Reapplying text values after swap...');
              let reapplyIndex = 0;
              async function reapplyTextValues(node: SceneNode) {
                if (node.type === 'TEXT') {
                  const uniqueKey = `text_${reapplyIndex++}`;
                  const captured = textValues.get(uniqueKey);
                  const textNode = node as TextNode;
                  
                  // CRITICAL: Check if this text node is controlled by a component property.
                  // Only skip restoration IF the controlling property was explicitly mapped/set by us.
                  // If it's controlled by an unmapped property, we SHOULD overwrite it with the legacy text logic
                  // so that unmapped "content" carries over.
                  let shouldSkip = false;
                  if (textNode.componentPropertyReferences && textNode.componentPropertyReferences.characters) {
                      const propId = textNode.componentPropertyReferences.characters;
                      
                      // Check if this property ID was mapped (either directly or via name matching)
                      // We need to check finalPropsMap keys against the propId
                      if (finalPropsMap[propId] !== undefined) {
                          shouldSkip = true;
                          console.log(`  ℹ️ Skipping restoration for text node "${textNode.name}" (Active Mapped Property: ${propId})`);
                      } else {
                          // Try to find if we mapped this property but with a different ID (unlikely if setProperties worked, but possible)
                          // Also check if the *value* of the property matches our mapped value, implying it's already correct?
                          // console.log(`  [Debug] Text node "${textNode.name}" is bound to ${propId}, but that property was NOT in our map. We will overwrite text.`);
                      }
                  }
                  
                  if (shouldSkip) {
                        // FORCE UPDATE: Ensure the visual text matches the property value we set
                        if (textNode.componentPropertyReferences?.characters) {
                            const propId = textNode.componentPropertyReferences.characters;
                            const val = finalPropsMap[propId];
                            if (val !== undefined) {
                                try {
                                    // Load font to ensure we can edit
                                    if (textNode.fontName && typeof textNode.fontName === 'object') {
                                        await figma.loadFontAsync(textNode.fontName as FontName);
                                    }
                                    const newVal = String(val);
                                    if (textNode.characters !== newVal) {
                                        textNode.characters = newVal;
                                        console.log(`    ✅ [Force] Updated text node "${textNode.name}" to match property: "${newVal}"`);
                                    }
                                } catch (e) {
                                    console.warn(`    ⚠️ Failed to force update text node "${textNode.name}":`, e);
                                }
                            }
                        }
                        return;
                  }

                  if (captured) {
                    try {
                      // Load the font before setting text
                      if (captured.fontName && typeof captured.fontName === 'object') {
                        await figma.loadFontAsync(captured.fontName);
                      }
                      textNode.characters = captured.text;
                      // console.log(`  ✅ Reapplied text #${reapplyIndex - 1}`);
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
              // console.log('✅ Text reapplication complete');
              

              // Reapply captured nested instance property overrides
              // We perform this even if preserveStyle is false, as nested instances are Structure/Content, not Style.
              if (nestedInstanceOverrides.size > 0) {
                // console.log(`🔄 Reapplying nested instance overrides after swap...`);
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
                        
                        // Check if this nested instance is bound to a top-level property (Swap Property)
                        // Or if its properties are bound to top-level properties
                        // If so, we might want to skip overwriting it?
                        
                        // [Fix] If we have explicit mapped values for this component, we should be careful about
                        // blindly reapplying overrides found on the old instance structure.
                        
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
                          
                          // [Critical Fix] If this nested property is actually controlled by a top-level property
                          // that we JUST mapped, we must NOT overwrite it with the old captured value.
                          let isControlledByMappedProp = false;
                          if (fullPropName) {
                              // Cast to any to allow checking generic property names against the references map
                              // (Standard types only include visible, characters, mainComponent)
                              const refs = instNode.componentPropertyReferences as any;
                              const ref = refs?.[fullPropName.split('#')[0]];
                              
                              if (ref && finalPropsMap[ref]) {
                                  isControlledByMappedProp = true;
                                  console.log(`    ℹ️ Skipping nested override for '${fullPropName}' (Controlled by mapped prop '${ref}')`);
                              }
                          }
                          
                          if (fullPropName && !isControlledByMappedProp) {
                            newProps[fullPropName] = value;
                          }
                        }
                        
                        if (Object.keys(newProps).length > 0) {
                          instNode.setProperties(newProps);
                          // console.log(`  ✅ Applied ${Object.keys(newProps).length} properties to nested instance #${reapplyOverrideIndex}`);
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
                // console.log('✅ Nested instance override reapplication complete');
              }

              // DOUBLE CHECK: Re-apply properties one last time to ensure they weren't overwritten by overrides
              if (propertiesApplied && Object.keys(finalPropsMap).length > 0) {
                  try {
                       instanceNode.setProperties(finalPropsMap);
                       console.log('  🛡️  Re-enforced mapped properties (after overrides).');
                  } catch (e) {
                       // ignore
                  }
              }
              
              // Reapply captured layout overrides
              if (nestedInstanceLayouts.size > 0) {
                // console.log('🔄 Reapplying layout overrides after swap...');
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
                          } catch (e) {console.warn(e);}
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
                    // console.log('✅ Layout override reapplication complete');
                  }
              
              // Restore position only
              try {
                instanceNode.x = x;
                instanceNode.y = y;
                instanceNode.rotation = rotation;
              } catch (e) { console.warn('Could not restore position', e); }
              
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
async function runDiagnostics() {
  console.log('🔍 Running Diagnostics...');
  
  // 1. Check API Version
  console.log(`  - API Version: ${figma.apiVersion}`);
  
  // 2. Check User (Skipped to avoid permission error)
  // if (figma.currentUser) {
  //     console.log(`  - User: ${figma.currentUser.name} (ID: ${figma.currentUser.id})`);
  // } else {
  //     console.log(`  - User: Unknown (figma.currentUser is null)`);
  // }
  
  // 3. Check Team Library Capabilities
  console.log(`  - figma.teamLibrary keys: ${Object.keys(figma.teamLibrary).join(', ')}`);
  
  // 4. Check Local Variables
  try {
      const localVars = await figma.variables.getLocalVariablesAsync();
      console.log(`  - Local Variables: ${localVars.length}`);
  } catch (e) {
      console.log(`  - Local Variables Check Failed: ${e}`);
  }

  // 5. Check Available Collections (Raw)
  try {
      const cols = await figma.teamLibrary.getAvailableLibraryVariableCollectionsAsync();
      console.log(`  - Available Collections: ${cols.length}`);
      cols.forEach(c => console.log(`    * ${c.name} (${c.libraryName}) - Key: ${c.key}`));
  } catch (e) {
      console.log(`  - Available Collections Check Failed: ${e}`);
  }

  // 6. PROBE: Try to import a known variable key from "Monkey"
  // Key from previous logs: dc69fade742a1338bc34ec90e4081f924f45fbbb
  const probeKey = 'dc69fade742a1338bc34ec90e4081f924f45fbbb';
  console.log(`  - PROBE: Attempting to import known variable key: ${probeKey}`);
  try {
      const v = await figma.variables.importVariableByKeyAsync(probeKey);
      if (v) {
          console.log(`    ✅ SUCCESS! Imported variable: ${v.name} from ${v.variableCollectionId}`);
          console.log(`    ℹ️ CONCLUSION: The library IS accessible, but getAvailableLibraryVariableCollectionsAsync is failing.`);
      } else {
          console.log(`    ❌ FAILED: importVariableByKeyAsync returned null.`);
          console.log(`    ℹ️ CONCLUSION: The library is NOT accessible to this file.`);
      }
  } catch (e) {
      console.log(`    ❌ FAILED: importVariableByKeyAsync threw error: ${e}`);
  }
  
  figma.notify('Diagnostics complete. Check console.');
}

// Helper: Manual Base64 Encoder (fallback for environments without btoa)
function bufferToBase64(buffer: Uint8Array): string {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
    const bytes = new Uint8Array(buffer);
    let i = 0;
    let len = bytes.length;
    let base64 = '';

    while (i < len) {
        const c1 = bytes[i++] & 0xff;
        if (i == len) {
            base64 += chars.charAt(c1 >> 2);
            base64 += chars.charAt((c1 & 0x3) << 4);
            base64 += '==';
            break;
        }
        const c2 = bytes[i++];
        if (i == len) {
            base64 += chars.charAt(c1 >> 2);
            base64 += chars.charAt(((c1 & 0x3) << 4) | ((c2 & 0xf0) >> 4));
            base64 += chars.charAt((c2 & 0xf) << 2);
            base64 += '=';
            break;
        }
        const c3 = bytes[i++];
        base64 += chars.charAt(c1 >> 2);
        base64 += chars.charAt(((c1 & 0x3) << 4) | ((c2 & 0xf0) >> 4));
        base64 += chars.charAt(((c2 & 0xf) << 2) | ((c3 & 0xc0) >> 6));
        base64 += chars.charAt(c3 & 0x3f);
    }
    return base64;
}

// Store context for live updates
let lastPropContext: {
    sourceId: string;
    targetKey: string;
    targetName: string;
    targetLibraryName?: string;
    sourceLibraryName?: string;
} | null = null;

// Listen for property changes on the source instance to update the UI dynamically
// figma.on('documentchange', (event) => {
//     if (!lastPropContext) return;
    
//     // Check if the source node had a property change
//     const changed = event.documentChanges.find(c => 
//         c.type === 'PROPERTY_CHANGE' && c.id === lastPropContext!.sourceId
//     );
    
//     if (changed) {
//         console.log('🔄 Source instance properties changed, refreshing mapping view...');
//         handleGetComponentProperties(
//             lastPropContext.sourceId,
//             lastPropContext.targetKey,
//             lastPropContext.targetName,
//             lastPropContext.targetLibraryName,
//             lastPropContext.sourceLibraryName
//         );
//     }
// });

async function handleGetComponentProperties(sourceId: string, targetKey: string, targetName: string, targetLibraryName?: string, sourceLibraryName?: string) {
    // Update context for live updates
    lastPropContext = { sourceId, targetKey, targetName, targetLibraryName, sourceLibraryName };

    try {
        console.log(`Getting properties for sourceId: ${sourceId}, targetKey: ${targetKey}, Lib: ${targetLibraryName}`);
        
        // 1. Get Source Definitions & Values
        let sourceDefinitions = {};
        let sourcePropertyValues: any = {};
        let sourceComponentName = "Source Component";
        const sourceNode = await figma.getNodeByIdAsync(sourceId);
        
        if (sourceNode) {
             if (sourceNode.type === 'INSTANCE') {
                try {
                    // Extract values from Instance (overrides)
                    sourcePropertyValues = sourceNode.componentProperties;
                    
                    const main = await sourceNode.getMainComponentAsync();
                    if (main) {
                        if (main.parent && main.parent.type === 'COMPONENT_SET') {
                            sourceDefinitions = main.parent.componentPropertyDefinitions;
                            sourceComponentName = main.parent.name;
                        } else {
                            sourceDefinitions = main.componentPropertyDefinitions;
                            sourceComponentName = main.name;
                        }
                    }
                } catch (e) { console.warn("Could not get main component", e); }
             } else if (sourceNode.type === 'COMPONENT') {
                  // Extract values from Component (Variant properties)
                  // For a Component (Variant), componentProperties is not directly populated like an Instance.
                  // We need to construct it from value-based properties (if explicit) or parse name (legacy)?
                  // Actually, ComponentNode normally doesn't have .componentProperties with values.
                  // It represents a specific combination of properties.
                  // Valid properties are in sourceNode.variantProperties or we have to derive them?
                  // Figma API: ComponentNode.variantProperties returns { [property: string]: string } | null
                  
                  if (sourceNode.variantProperties) {
                      sourcePropertyValues = {};
                      // Map variantProperties (simple key-value) to componentProperties format (key -> {value, type})
                      // But we need the full property ID (Name#ID) to match definitions.
                      // sourceNode.variantProperties only gives "Name" : "Value"
                      
                      const defs = sourceNode.parent && sourceNode.parent.type === 'COMPONENT_SET' 
                          ? sourceNode.parent.componentPropertyDefinitions 
                          : sourceNode.componentPropertyDefinitions;
                          
                      for (const [propName, propVal] of Object.entries(sourceNode.variantProperties)) {
                          // Find matching definition to get the full key (Name#ID)
                          const fullKey = Object.keys(defs).find(key => key.startsWith(propName + '#') || key === propName);
                          if (fullKey) {
                              sourcePropertyValues[fullKey] = { 
                                  value: propVal, 
                                  type: defs[fullKey].type 
                              };
                          }
                      }
                  } else {
                     // Fallback if no variantProperties (e.g. simple component)
                     // Use default values from definitions
                     const defs = sourceNode.componentPropertyDefinitions;
                     if (defs) {
                         for (const [key, def] of Object.entries(defs)) {
                             sourcePropertyValues[key] = {
                                 value: def.defaultValue,
                                 type: def.type
                             };
                         }
                     }
                  }
                  
                  console.log('Component Variant Values:', sourcePropertyValues);

                  // If it's a variant, get definitions from the parent ComponentSet
                  if (sourceNode.parent && sourceNode.parent.type === 'COMPONENT_SET') {
                      sourceDefinitions = sourceNode.parent.componentPropertyDefinitions;
                      sourceComponentName = sourceNode.parent.name;
                  } else {
                      sourceDefinitions = sourceNode.componentPropertyDefinitions;
                      sourceComponentName = sourceNode.name;
                  }
             } else if (sourceNode.type === 'COMPONENT_SET') {
                  sourceDefinitions = sourceNode.componentPropertyDefinitions;
                  sourceComponentName = sourceNode.name;
                  // Component Sets don't have values themselves (defaults are in definitions)
             }
        }
        
        // Resolve Instance Swap IDs to Names for display
        for (const key in sourcePropertyValues) {
             const prop = sourcePropertyValues[key];
             if (prop.type === 'INSTANCE_SWAP' && typeof prop.value === 'string') {
                 try {
                     const swapId = prop.value;
                     if (swapId && swapId.length > 0) {
                         const swappedNode = await figma.getNodeByIdAsync(swapId);
                         if (swappedNode) {
                             sourcePropertyValues[key] = {
                                 ...prop,
                                 value: swappedNode.name // Replace ID with Name for UI display
                             };
                         }
                     }
                 } catch (e) {
                     console.warn('Failed to resolve instance swap name:', e);
                 }
             }
        }

        // 2. Get Target Definitions
        let targetDefinitions = {};
        if (targetKey) {
            try {
                // Try importing as component
                let importedParam: any;
                try {
                    importedParam = await figma.importComponentByKeyAsync(targetKey);
                } catch (e) {
                    // Try as set
                    importedParam = await figma.importComponentSetByKeyAsync(targetKey);
                }
                
                if (importedParam) {
                    if (importedParam.type === 'COMPONENT_SET') {
                        targetDefinitions = importedParam.componentPropertyDefinitions;
                    } else if (importedParam.type === 'COMPONENT') {
                        if (importedParam.parent && importedParam.parent.type === 'COMPONENT_SET') {
                             targetDefinitions = importedParam.parent.componentPropertyDefinitions;
                        } else {
                             targetDefinitions = importedParam.componentPropertyDefinitions;
                        }
                    }
                }
            } catch (err) {
                console.error("Error importing target component:", err);
            }
        }

        figma.ui.postMessage({
            type: 'SHOW_PROPERTY_MAPPING_VIEW',
            sourceId,
            targetName,
            targetLibraryName,
            sourceLibraryName,
            sourceComponentName,
            sourceDefinitions,
            sourcePropertyValues,
            targetDefinitions
        });
        
    } catch (err: any) {
        console.error("Error in handleGetComponentProperties:", err);
        figma.notify("Error loading properties: " + err.message);
    }
}
