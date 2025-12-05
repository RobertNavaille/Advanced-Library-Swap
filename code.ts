// Main Figma Swap Library Plugin (clean template)
import { COMPONENT_KEY_MAPPING as DEFAULT_COMPONENT_KEY_MAPPING, STYLE_KEY_MAPPING as DEFAULT_STYLE_KEY_MAPPING, VARIABLE_ID_MAPPING as DEFAULT_VARIABLE_ID_MAPPING, VARIABLE_KEY_MAPPING as DEFAULT_VARIABLE_KEY_MAPPING, LIBRARY_THUMBNAILS as DEFAULT_LIBRARY_THUMBNAILS } from './keyMapping';
import { copyTextOverrides } from './swapUtils';

// JSONBin configuration
const JSONBIN_BIN_ID = '69324103d0ea881f4013bbac';
const JSONBIN_ACCESS_KEY = '$2a$10$LO/AIA/nruYOWBg5fzDPLOZI1Y8vv.Yucd1KsJBl6BquPRecLEf5C';
const JSONBIN_API_URL = `https://api.jsonbin.io/v3/b/${JSONBIN_BIN_ID}`;

// Global mapping variables - will be populated from JSONBin
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

// Fetch mappings from JSONBin
async function fetchMappingsFromJSONBin(): Promise<void> {
  try {
    console.log('📡 Fetching mappings from JSONBin...');
    const response = await fetch(JSONBIN_API_URL, {
      method: 'GET',
      headers: {
        'X-Access-Key': JSONBIN_ACCESS_KEY,
      }
    });
    
    if (!response.ok) {
      throw new Error(`JSONBin fetch failed: ${response.status} ${response.statusText}`);
    }
    
    const data = await response.json();
    const themes = data.record?.themes || data.themes;
    
    if (!themes) {
      throw new Error('No themes found in JSONBin response');
    }
    
    console.log('✅ Successfully fetched mappings from JSONBin');
    
    // Extract and build the mapping objects from themes
    COMPONENT_KEY_MAPPING = {};
    STYLE_KEY_MAPPING = {};
    VARIABLE_KEY_MAPPING = {};
    VARIABLE_ID_MAPPING = {};
    LIBRARY_THUMBNAILS = {};
    
    for (const [themeName, themeData]: [string, any] of Object.entries(themes)) {
      COMPONENT_KEY_MAPPING[themeName] = themeData.componentKeyMapping || {};
      STYLE_KEY_MAPPING[themeName] = themeData.styleKeyMapping || {};
      VARIABLE_KEY_MAPPING[themeName] = themeData.variableKeyMapping || {};
      VARIABLE_ID_MAPPING[themeName] = themeData.variableIdMapping || {};
      LIBRARY_THUMBNAILS[themeName] = themeData.thumbnail || '';
    }
    
    console.log('🎨 Loaded themes:', Object.keys(COMPONENT_KEY_MAPPING));
  } catch (error) {
    console.error('❌ Error fetching from JSONBin, using local defaults:', error);
    // Mappings remain as defaults from keyMapping.ts
  }
}

// Upload mappings to JSONBin
async function uploadToJSONBin(components: { name: string; key: string }[], styles: { name: string; key: string; type: string }[], variables: { name: string; id: string; key?: string }[]): Promise<void> {
  try {
    console.log('📤 Uploading to JSONBin...');
    
    // Detect library name from file name
    const fileName = figma.root.name.toLowerCase();
    let detectedLibrary = 'Custom';
    if (fileName.includes('shark')) {
      detectedLibrary = 'Shark';
    } else if (fileName.includes('monkey')) {
      detectedLibrary = 'Monkey';
    }
    console.log(`🏷️ Detected library from file name: ${detectedLibrary}`);
    
    // Build component mapping (name -> key)
    const componentMapping: Record<string, string> = {};
    components.forEach(comp => {
      componentMapping[comp.name] = comp.key;
    });
    
    // Build style mapping (name -> key)
    const styleMapping: Record<string, string> = {};
    styles.forEach(style => {
      styleMapping[style.name] = style.key;
    });
    
    // Build variable mappings (use key if available, fall back to id)
    const variableKeyMapping: Record<string, string> = {};
    const variableIdMapping: Record<string, string> = {};
    variables.forEach(variable => {
      variableKeyMapping[variable.name] = variable.key || variable.id;
      variableIdMapping[variable.name] = variable.id;
    });
    
    // Create updated mappings structure, updating the detected library
    const updatedMappings = {
      themes: {
        ...COMPONENT_KEY_MAPPING,
        [detectedLibrary]: {
          componentKeyMapping: componentMapping,
          styleKeyMapping: styleMapping,
          variableKeyMapping: variableKeyMapping,
          variableIdMapping: variableIdMapping,
          thumbnail: LIBRARY_THUMBNAILS[detectedLibrary] || ''
        }
      }
    };
    
    const response = await fetch(JSONBIN_API_URL, {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/json',
        'X-Access-Key': JSONBIN_ACCESS_KEY,
      },
      body: JSON.stringify(updatedMappings)
    });
    
    if (!response.ok) {
      throw new Error(`JSONBin upload failed: ${response.status} ${response.statusText}`);
    }
    
    console.log('✅ Successfully uploaded mappings to JSONBin');
    console.log(`🎨 Updated ${detectedLibrary} theme with synced components and styles`);
  } catch (error) {
    console.error('❌ Error uploading to JSONBin:', error);
    console.log('⚠️ Mappings were not uploaded, but sync is complete.');
  }
}

// Store the scanned frame for later use during swaps
let scannedFrame: FrameNode | null = null;

// Show the UI
figma.showUI(__html__, { width: 480, height: 500, themeColors: true });

// Fetch mappings from JSONBin on plugin load
fetchMappingsFromJSONBin();

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
    
    // Extract unique libraries and build library metadata with file IDs
    const libraryMap = new Map<string, { fileId?: string; components: number; tokens: number }>();
    
    // Add components to library map
    components.forEach(c => {
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
    tokens.forEach(t => {
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
    
    figma.ui.postMessage({ type: 'SCAN_ALL_RESULT', ok: true, data: { components, tokens, libraries } });
  } catch (error) {
    console.error('❌ Scan error:', error);
    figma.ui.postMessage({ type: 'SCAN_ALL_RESULT', ok: false, error: `Scan failed: ${error instanceof Error ? error.message : 'Unknown error'}` });
  }
}

// Recursively scan nodes for components and design tokens
async function scanNodeForAssets(node: SceneNode, components: ComponentInfo[], tokens: TokenInfo[]): Promise<void> {
  if (node.type === 'INSTANCE') {
    try {
      const component = await node.getMainComponentAsync();
      if (component) {
        let foundLibrary: string | null = null;
        let foundName: string | null = null;
        let parentName: string | null = null;
        
        // Extract file ID from component key (format: fileId/pageId/componentId)
        const keyParts = component.key.split('/');
        const libraryFileId = keyParts[0] || undefined;
        const componentId = keyParts[keyParts.length - 1] || component.key; // Get the last part (component ID)
        
        // Try to match the component using the full key first, then try with extracted component ID
        for (const libName of Object.keys(COMPONENT_KEY_MAPPING)) {
          for (const compName of Object.keys(COMPONENT_KEY_MAPPING[libName])) {
            const mappedKey = COMPONENT_KEY_MAPPING[libName][compName];
            if (mappedKey === component.key || mappedKey === componentId) {
              foundLibrary = libName;
              foundName = compName;
              if (component.parent && component.parent.type === 'COMPONENT_SET') {
                parentName = component.parent.name;
              } else {
                parentName = foundName;
              }
              break;
            }
          }
          if (foundLibrary) break;
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

// Get target color for UI display
async function getTargetColor(tokenId: string, styleName: string, sourceLibrary: string, targetLibrary: string): Promise<void> {
  console.log(`🎨 Getting target color for: ${styleName}, from ${sourceLibrary} to ${targetLibrary}`);
  console.log(`📦 tokenId: ${tokenId}`);
  
  function normalizeLibraryName(name: string): string {
    if (name.toLowerCase().includes('shark')) return 'Shark';
    if (name.toLowerCase().includes('monkey')) return 'Monkey';
    return name;
  }
  
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
  await uploadToJSONBin(components, styles, variables);
  
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

  // Get the source style key - we need this to identify which styles to replace
  const sourceStyleKey = STYLE_KEY_MAPPING[sourceLibrary]?.[styleName];
  
  if (!sourceStyleKey) {
    console.warn(`⚠️ Source style mapping not found for '${styleName}' in ${sourceLibrary}`);
    return 0;
  }

  // Normalize target library name
  function normalizeLibraryName(name: string): string {
    if (name.toLowerCase().includes('shark')) return 'Shark';
    if (name.toLowerCase().includes('monkey')) return 'Monkey';
    return name;
  }
  const normalizedTargetLibrary = normalizeLibraryName(targetLibrary);

  // Try to get the target variable KEY for this style name
  const targetVariableKey = VARIABLE_KEY_MAPPING[normalizedTargetLibrary]?.[styleName];
  const targetVariableId = VARIABLE_ID_MAPPING[normalizedTargetLibrary]?.[styleName];

  console.log(`  📋 Processing node: ${node.name} (type: ${node.type}) looking for style: ${styleName}`);

  // Process fill styles on any node that has them (including FRAME)
  if ('fillStyleId' in node && node.fillStyleId && typeof node.fillStyleId === 'string') {
    try {
      const currentStyle = await figma.getStyleByIdAsync(node.fillStyleId);
      if (currentStyle && currentStyle.key === sourceStyleKey) {
        console.log(`  ✅ Found ${styleName} style on ${node.type}`);
        
        // If we have a target variable KEY or ID, bind to it
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
      swapCount += await swapStylesInNode(child, styleName, sourceLibrary, normalizedTargetLibrary);
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

  function normalizeLibraryName(name: string): string {
    if (name.toLowerCase().includes('shark')) return 'Shark';
    if (name.toLowerCase().includes('monkey')) return 'Monkey';
    return name;
  }
  const normalizedSourceLibrary = normalizeLibraryName(sourceLibrary);
  const normalizedTargetLibrary = normalizeLibraryName(targetLibrary);

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
          const instances = await findInstancesByNameAsync(node, comp.name, normalizedSourceLibrary);
          totalInstancesFound += instances.length;
          if (instances.length === 0) {
            continue;  // Don't count as error yet, check other nodes
          }
          foundInstancesForThisComponent = true;
          for (const instance of instances) {
            const targetKey = COMPONENT_KEY_MAPPING[normalizedTargetLibrary]?.[comp.name];
            if (!targetKey) {
              errorDetails.push(`No mapping for component '${comp.name}' in target library '${normalizedTargetLibrary}'`);
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
        const styleSwaps = await swapStylesInNode(node, style.name, normalizedSourceLibrary, normalizedTargetLibrary);
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