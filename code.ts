// Main Figma Swap Library Plugin (clean template)
import { COMPONENT_KEY_MAPPING, STYLE_KEY_MAPPING } from './keyMapping';
import { copyTextOverrides } from './swapUtils';

// Interface definitions
interface ComponentInfo {
  id: string;
  name: string;  // The variant name (mapping key) - used for swapping
  displayName: string;  // The parent component name - used for UI display
  library: string;
  remote: boolean;
  parentName: string;  // Kept for backward compatibility
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
    figma.ui.postMessage({ type: 'SCAN_ALL_RESULT', ok: true, data: { components, tokens, libraries: [] } });
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
        for (const libName of Object.keys(COMPONENT_KEY_MAPPING)) {
          for (const compName of Object.keys(COMPONENT_KEY_MAPPING[libName])) {
            if (COMPONENT_KEY_MAPPING[libName][compName] === component.key) {
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
              parentName: parentComponentName 
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
    return `rgb(${r}, ${g}, ${b})`;
  }
  return '#000000';
}

// Get target color for UI display
async function getTargetColor(tokenId: string, styleName: string, sourceLibrary: string, targetLibrary: string): Promise<void> {
  console.log(`🎨 Getting target color for: ${styleName}, from ${sourceLibrary} to ${targetLibrary}`);
  
  function normalizeLibraryName(name: string): string {
    if (name.toLowerCase().includes('shark')) return 'Shark';
    if (name.toLowerCase().includes('monkey')) return 'Monkey';
    return name;
  }
  
  const normalizedTargetLibrary = normalizeLibraryName(targetLibrary);
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
  
  // Load all pages first
  await figma.loadAllPagesAsync();
  
  // Get all local components from current page only (more practical)
  const localComponents = figma.currentPage.findAll(node => node.type === 'COMPONENT') as ComponentNode[];
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
  
  // Output to console for easy copying
  console.log('\n📦 COMPONENT KEYS:');
  components.forEach(comp => {
    console.log(`  '${comp.name}': '${comp.key}',`);
  });
  
  console.log('\n🎨 STYLE KEYS:');
  styles.forEach(style => {
    console.log(`  '${style.name}': '${style.key}', // ${style.type}`);
  });
  
  figma.notify(`Found ${components.length} components and ${styles.length} styles. Check console for keys.`);
  figma.ui.postMessage({ 
    type: 'SYNC_COMPLETE', 
    components: components.length, 
    styles: styles.length 
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

  // Get the source and target style keys
  const sourceStyleKey = STYLE_KEY_MAPPING[sourceLibrary]?.[styleName];
  const targetStyleKey = STYLE_KEY_MAPPING[targetLibrary]?.[styleName];

  if (!sourceStyleKey || !targetStyleKey) {
    console.warn(`Style mapping not found for '${styleName}' between ${sourceLibrary} and ${targetLibrary}`);
    return 0;
  }

  // Check if this node uses the source style
  try {
    // Handle fill/paint styles
    if ('fillStyleId' in node && node.fillStyleId && typeof node.fillStyleId === 'string') {
      const currentStyle = await figma.getStyleByIdAsync(node.fillStyleId);
      if (currentStyle && currentStyle.key === sourceStyleKey) {
        // Import and apply the target style
        const targetStyle = await figma.importStyleByKeyAsync(targetStyleKey);
        if (targetStyle && targetStyle.type === 'PAINT') {
          await node.setFillStyleIdAsync(targetStyle.id);
          swapCount++;
          console.log(`✅ Swapped fill style '${styleName}' on ${node.type}`);
        }
      }
    }

    // Handle stroke styles
    if ('strokeStyleId' in node && node.strokeStyleId && typeof node.strokeStyleId === 'string') {
      const currentStyle = await figma.getStyleByIdAsync(node.strokeStyleId);
      if (currentStyle && currentStyle.key === sourceStyleKey) {
        const targetStyle = await figma.importStyleByKeyAsync(targetStyleKey);
        if (targetStyle && targetStyle.type === 'PAINT') {
          await node.setStrokeStyleIdAsync(targetStyle.id);
          swapCount++;
          console.log(`✅ Swapped stroke style '${styleName}' on ${node.type}`);
        }
      }
    }

    // Handle text styles
    if ('textStyleId' in node && node.textStyleId && typeof node.textStyleId === 'string') {
      const currentStyle = await figma.getStyleByIdAsync(node.textStyleId);
      if (currentStyle && currentStyle.key === sourceStyleKey) {
        const targetStyle = await figma.importStyleByKeyAsync(targetStyleKey);
        if (targetStyle && targetStyle.type === 'TEXT') {
          await node.setTextStyleIdAsync(targetStyle.id);
          swapCount++;
          console.log(`✅ Swapped text style '${styleName}' on ${node.type}`);
        }
      }
    }

    // Handle effect styles
    if ('effectStyleId' in node && node.effectStyleId && typeof node.effectStyleId === 'string') {
      const currentStyle = await figma.getStyleByIdAsync(node.effectStyleId);
      if (currentStyle && currentStyle.key === sourceStyleKey) {
        const targetStyle = await figma.importStyleByKeyAsync(targetStyleKey);
        if (targetStyle && targetStyle.type === 'EFFECT') {
          await node.setEffectStyleIdAsync(targetStyle.id);
          swapCount++;
          console.log(`✅ Swapped effect style '${styleName}' on ${node.type}`);
        }
      }
    }
  } catch (err) {
    console.warn(`Error swapping style on node:`, err);
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
      
      for (const node of nodesToProcess) {
        if (node.type === 'FRAME' || node.type === 'GROUP') {
          const instances = await findInstancesByNameAsync(node, comp.name, normalizedSourceLibrary);
          totalInstancesFound += instances.length;
          if (instances.length === 0) {
            errorDetails.push(`No matching instances found for component '${comp.name}' in selection.`);
            errorCount++;
            continue;
          }
          for (const instance of instances) {
            const targetKey = COMPONENT_KEY_MAPPING[normalizedTargetLibrary]?.[comp.name];
            if (!targetKey) {
              errorDetails.push(`No mapping for component '${comp.name}' in target library '${normalizedTargetLibrary}'`);
              errorCount++;
              continue;
            }
            try {
              const importedComponent = await figma.importComponentByKeyAsync(targetKey);
              
              // Store position only (not size, which becomes read-only after swap)
              const instanceNode = instance as InstanceNode;
              const x = instanceNode.x;
              const y = instanceNode.y;
              const rotation = instanceNode.rotation;
              
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
        if (node.type === 'FRAME' || node.type === 'GROUP') {
          const styleSwaps = await swapStylesInNode(node, style.name, normalizedSourceLibrary, normalizedTargetLibrary);
          styleSwapCount += styleSwaps;
        }
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