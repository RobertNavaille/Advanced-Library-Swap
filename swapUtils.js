// swapUtils.ts
// Utility functions for swapping and copying properties between Figma instances
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
// Use Figma's global InstanceNode type
export function copyTextOverrides(sourceInstance, targetInstance) {
    return __awaiter(this, void 0, void 0, function* () {
        // Helper to copy variant properties
        function copyVariantProps(source, target) {
            if (source.variantProperties && target.variantProperties) {
                const variantProps = {};
                for (const key of Object.keys(source.variantProperties)) {
                    if (target.variantProperties[key] !== undefined) {
                        variantProps[key] = String(source.variantProperties[key]);
                    }
                }
                if (Object.keys(variantProps).length > 0) {
                    try {
                        target.setProperties(variantProps);
                    }
                    catch (err) {
                        console.warn('Could not set variant properties:', err);
                    }
                }
            }
        }
        // Helper to preserve instance swap for nested instances
        function preserveInstanceSwap(source, target) {
            return __awaiter(this, void 0, void 0, function* () {
                let sourceMain = null;
                let targetMain = null;
                try {
                    sourceMain = yield source.getMainComponentAsync();
                }
                catch (err) {
                    console.warn(`preserveInstanceSwap: Failed to get source main component.`, err);
                    return;
                }
                try {
                    targetMain = yield target.getMainComponentAsync();
                }
                catch (err) {
                    console.warn(`preserveInstanceSwap: Failed to get target main component.`, err);
                    return;
                }
                if (!sourceMain || !targetMain)
                    return;
                if (sourceMain.key !== targetMain.key) {
                    try {
                        yield target.swapComponent(sourceMain);
                    }
                    catch (err) {
                        console.warn(`preserveInstanceSwap: Failed to swap component.`, err);
                    }
                }
            });
        }
        // Helper to copy string/boolean properties by base name
        function copyPropsByBaseName(source, target) {
            if (source.componentProperties && target.componentProperties) {
                const newProps = {};
                const targetTextProps = {};
                const targetBoolProps = {};
                for (const targetPropName of Object.keys(target.componentProperties)) {
                    const tProp = target.componentProperties[targetPropName];
                    const base = targetPropName.split('#')[0];
                    if (tProp && typeof tProp === 'object' && tProp.type === 'TEXT') {
                        targetTextProps[base] = targetPropName;
                    }
                    else if (tProp && typeof tProp === 'object' && tProp.type === 'BOOLEAN') {
                        targetBoolProps[base] = targetPropName;
                    }
                }
                for (const sourcePropName of Object.keys(source.componentProperties)) {
                    const sProp = source.componentProperties[sourcePropName];
                    const base = sourcePropName.split('#')[0];
                    if (sProp && typeof sProp === 'object' && sProp.type === 'TEXT' && targetTextProps[base]) {
                        newProps[targetTextProps[base]] = sProp.value;
                    }
                    else if (sProp && typeof sProp === 'object' && sProp.type === 'BOOLEAN' && targetBoolProps[base]) {
                        newProps[targetBoolProps[base]] = sProp.value;
                    }
                }
                if (Object.keys(newProps).length > 0) {
                    try {
                        target.setProperties(newProps);
                    }
                    catch (err) {
                        console.warn('Could not set string/boolean properties:', err);
                    }
                }
            }
        }
        // Run all helpers
        copyVariantProps(sourceInstance, targetInstance);
        yield preserveInstanceSwap(sourceInstance, targetInstance);
        copyPropsByBaseName(sourceInstance, targetInstance);
    });
}
