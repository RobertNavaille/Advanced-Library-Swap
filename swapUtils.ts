// swapUtils.ts
// Utility functions for swapping and copying properties between Figma instances

// Use Figma's global InstanceNode type

export async function copyTextOverrides(sourceInstance: InstanceNode, targetInstance: InstanceNode): Promise<void> {
  // Helper to copy variant properties
  function copyVariantProps(source: InstanceNode, target: InstanceNode) {
    if (source.variantProperties && target.variantProperties) {
      const variantProps: {[key: string]: string} = {};
      for (const key of Object.keys(source.variantProperties)) {
        if (target.variantProperties[key] !== undefined) {
          variantProps[key] = String(source.variantProperties[key]);
        }
      }
      if (Object.keys(variantProps).length > 0) {
        try {
          target.setProperties(variantProps);
        } catch (err) {
          console.warn('Could not set variant properties:', err);
        }
      }
    }
  }
  // Helper to preserve instance swap for nested instances
  async function preserveInstanceSwap(source: InstanceNode, target: InstanceNode) {
    let sourceMain = null;
    let targetMain = null;
    try {
      sourceMain = await source.getMainComponentAsync();
    } catch (err) {
      console.warn(`preserveInstanceSwap: Failed to get source main component.`, err);
      return;
    }
    try {
      targetMain = await target.getMainComponentAsync();
    } catch (err) {
      console.warn(`preserveInstanceSwap: Failed to get target main component.`, err);
      return;
    }
    if (!sourceMain || !targetMain) return;
    if (sourceMain.key !== targetMain.key) {
      try {
        await target.swapComponent(sourceMain);
      } catch (err) {
        console.warn(`preserveInstanceSwap: Failed to swap component.`, err);
      }
    }
  }
  // Helper to copy string/boolean properties by base name
  function copyPropsByBaseName(source: InstanceNode, target: InstanceNode) {
    if (source.componentProperties && target.componentProperties) {
      const newProps: {[key: string]: any} = {};
      const targetTextProps: {[base: string]: string} = {};
      const targetBoolProps: {[base: string]: string} = {};
      for (const targetPropName of Object.keys(target.componentProperties)) {
        const tProp = target.componentProperties[targetPropName];
        const base = targetPropName.split('#')[0];
        if (tProp && typeof tProp === 'object' && tProp.type === 'TEXT') {
          targetTextProps[base] = targetPropName;
        } else if (tProp && typeof tProp === 'object' && tProp.type === 'BOOLEAN') {
          targetBoolProps[base] = targetPropName;
        }
      }
      for (const sourcePropName of Object.keys(source.componentProperties)) {
        const sProp = source.componentProperties[sourcePropName];
        const base = sourcePropName.split('#')[0];
        if (sProp && typeof sProp === 'object' && sProp.type === 'TEXT' && targetTextProps[base]) {
          newProps[targetTextProps[base]] = sProp.value;
        } else if (sProp && typeof sProp === 'object' && sProp.type === 'BOOLEAN' && targetBoolProps[base]) {
          newProps[targetBoolProps[base]] = sProp.value;
        }
      }
      if (Object.keys(newProps).length > 0) {
        try {
          target.setProperties(newProps);
        } catch (err) {
          console.warn('Could not set string/boolean properties:', err);
        }
      }
    }
  }
  // Run all helpers
  copyVariantProps(sourceInstance, targetInstance);
  await preserveInstanceSwap(sourceInstance, targetInstance);
  copyPropsByBaseName(sourceInstance, targetInstance);
}
