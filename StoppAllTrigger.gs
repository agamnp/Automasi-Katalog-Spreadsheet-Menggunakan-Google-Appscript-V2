function stopAllBatches() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  PropertiesService.getScriptProperties().deleteAllProperties();
  Logger.log("🛑 Semua trigger & properti batch dihapus total.");
}