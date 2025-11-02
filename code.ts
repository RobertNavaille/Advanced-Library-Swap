// Main Figma Swap Library Plugin (clean template)
import { COMPONENT_KEY_MAPPING, STYLE_KEY_MAPPING } from './keyMapping';
import { copyTextOverrides } from './swapUtils';

// Interface definitions
interface ComponentInfo {
  id: string;
  name: string;
  library: string;
  remote: boolean;
  parentName: string;
}

interface TokenInfo {
  id: string;
  name: string;
  type: 'color' | 'typography' | 'spacing' | 'effect';
  value: string;
}

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
    const selection = figma.currentPage.selection;
    if (selection.length === 0) {
      figma.ui.postMessage({ type: 'SCAN_ALL_RESULT', ok: false, error: 'Please select at least one frame to scan' });
      return;
    }
    const components: ComponentInfo[] = [];
    const tokens: TokenInfo[] = [];
    for (const node of selection) {
      if (node.type === 'FRAME') {
        await scanNodeForAssets(node, components, tokens);
      }
    }
    figma.ui.postMessage({ type: 'SCAN_ALL_RESULT', ok: true, data: { components, tokens, libraries: [] } });
  } catch (error) {
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
          const variantName = foundName; // Always use the mapping key (variant)
          const parentComponentName = parentName || foundName;
          if (!variantName.startsWith('.') && !components.find(c => c.name === variantName && c.library === foundLibrary)) {
            components.push({ id: component.id, name: variantName, library: foundLibrary, remote: component.remote, parentName: parentComponentName });
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
  if ('fillStyleId' in node && node.fillStyleId && typeof node.fillStyleId === 'string') {
    try {
      const paintStyle = await figma.getStyleByIdAsync(node.fillStyleId);
      if (paintStyle && paintStyle.type === 'PAINT') {
        const tokenInfo: TokenInfo = { id: paintStyle.id, name: paintStyle.name, type: 'color', value: getColorValue(paintStyle.paints[0]) };
        if (!tokens.find(t => t.id === tokenInfo.id)) tokens.push(tokenInfo);
      }
    } catch {}
  }
  if ('textStyleId' in node && node.textStyleId && typeof node.textStyleId === 'string') {
    try {
      const textStyle = await figma.getStyleByIdAsync(node.textStyleId);
      if (textStyle && textStyle.type === 'TEXT') {
        const tokenInfo: TokenInfo = { id: textStyle.id, name: textStyle.name, type: 'typography', value: `${textStyle.fontSize}px ${textStyle.fontName?.family || 'Unknown'}` };
        if (!tokens.find(t => t.id === tokenInfo.id)) tokens.push(tokenInfo);
      }
    } catch {}
  }
  if ('effectStyleId' in node && node.effectStyleId && typeof node.effectStyleId === 'string') {
    try {
      const effectStyle = await figma.getStyleByIdAsync(node.effectStyleId);
      if (effectStyle && effectStyle.type === 'EFFECT') {
        const tokenInfo: TokenInfo = { id: effectStyle.id, name: effectStyle.name, type: 'effect', value: 'Effect style' };
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

// Library swap stub
function handleSwapLibrary(libraryId: string): void {
  figma.ui.postMessage({ type: 'swap-complete', message: 'Library swap completed successfully' });
}

// Helper: Recursively find all instances by name and library
async function findInstancesByNameAsync(node: SceneNode, name: string, library: string): Promise<InstanceNode[]> {
  let found: InstanceNode[] = [];
  if (node.type === 'INSTANCE') {
    try {
      const mainComponent = await node.getMainComponentAsync();
      if (mainComponent && mainComponent.name === name) {
        // Check if instance belongs to the source library
        for (const lib in COMPONENT_KEY_MAPPING) {
          if (lib === library && COMPONENT_KEY_MAPPING[lib][name] === mainComponent.key) {
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
      const selection = figma.currentPage.selection;
      for (const node of selection) {
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
              instance.swapComponent(importedComponent);
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

  if (totalInstancesFound === 0 && errorCount === 0) {
    figma.ui.postMessage({ type: 'swap-error', message: 'No matching component instances found in selection.', details: [] });
    console.error('Swap error: No matching component instances found in selection.');
    return;
  }

  if (errorCount > 0) {
    figma.ui.postMessage({ type: 'swap-error', message: `Swap completed with ${errorCount} errors.`, details: errorDetails });
    console.error('Swap error details:', errorDetails);
  } else {
    figma.ui.postMessage({ type: 'swap-complete', message: `Swap completed successfully! ${swapCount} instances swapped.` });
    figma.notify(`Swap completed successfully! ${swapCount} instances swapped.`);
  }
}

// ===== DELETE EVERYTHING BELOW THIS LINE! =====