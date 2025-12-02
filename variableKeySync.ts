// Run this in the Monkey library file to get variable keys
// Go to Monkey library → Plugins → Run command → Then paste this into the console

export async function syncVariableKeys() {
  const localVars = await figma.variables.getLocalVariablesAsync();
  
  console.log('🔑 VARIABLE KEYS FOR MONKEY LIBRARY:');
  console.log('=====================================');
  
  const mapping: Record<string, string> = {};
  
  for (const variable of localVars) {
    if (variable.resolvedType === 'COLOR') {
      console.log(`'${variable.name}': '${variable.key}',`);
      mapping[variable.name] = variable.key;
    }
  }
  
  console.log('\nCopy this into keyMapping.ts under VARIABLE_KEY_MAPPING:');
  console.log(JSON.stringify(mapping, null, 2));
  
  return mapping;
}

// Call this function
syncVariableKeys();
