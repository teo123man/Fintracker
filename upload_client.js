/* Shared by the welcome screen and the dashboard; injected by the local server. */
globalThis.uploadStatement = async function(file, token, fetcher = fetch) {
  if (!/\.(pdf|xlsx)$/i.test(file.name) || file.size > 40 * 1024 * 1024) {
    throw new Error('Dosya PDF veya XLSX biçiminde ve en fazla 40 MB olmalıdır.');
  }
  let bytes;
  try {
    // Read before the request: distinguish local file access from a lost server.
    // This also avoids deferred file-backed uploads if the source file changes.
    bytes = await file.arrayBuffer();
  } catch {
    throw new Error('Seçilen dosya okunamadı. Dosyayı bilgisayarınıza indirin, yeniden seçin ve tekrar deneyin.');
  }
  let response;
  try {
    response = await fetcher('/api/upload', {
      method: 'POST',
      headers: {'X-Token': token, 'X-Filename': encodeURIComponent(file.name),
                'Content-Type': 'application/octet-stream'},
      body: bytes
    });
  } catch {
    throw new Error('Dosya okunabildi ancak yerel uygulamaya aktarılamadı. Fintracker başlangıç penceresinin açık ve panelin hazır olduğundan emin olun; Paneli aç düğmesiyle yeniden açıp deneyin.');
  }
  let result;
  try { result = await response.json(); }
  catch { throw new Error('Uygulamadan geçerli yanıt alınamadı. Paneli başlangıç penceresinden yeniden açın.'); }
  if (response.status === 403) {
    throw new Error('Bu sekmenin oturumu sona ermiş. Paneli başlangıç penceresinden yeniden açıp dosyayı tekrar seçin.');
  }
  if (!response.ok || !result.ok) throw new Error(result.log || 'Dosya kaydedilemedi. Veri klasörünün yazılabilir olduğunu kontrol edin.');
  return result;
};
