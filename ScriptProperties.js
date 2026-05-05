function listAllProperties() {
  const props = PropertiesService.getScriptProperties().getProperties();
  for (const [key, value] of Object.entries(props)) {
    console.log(`${key} = ${value}`);
  }
}
function setOneProperty() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("DEBUG", "FALSE");
}

function addOrUpdateProperty() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("CHUNK_OVERLAP_SECONDS", "3");
}

function deleteOneProperty() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("DEBUG_STRATEGIST=true");
}
