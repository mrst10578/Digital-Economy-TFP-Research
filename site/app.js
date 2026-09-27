const el = document.getElementById("buildStamp");
if (el) {
  const d = new Date();
  el.textContent = "Workspace build · " + d.toLocaleDateString("fa-IR");
}
