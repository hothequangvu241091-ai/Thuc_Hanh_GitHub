const state = {
  summary: null, catalog: [], category: "Tất cả", query: "", launcherQuery: "",
  packageMode: "core", activeLauncherGroup: "", historyLauncherId: "", historyLogName: ""
};
const titles = {
  overview: "Tổng quan hệ thống", launcher: "Trình chạy của bạn", voice: "Điều khiển giọng nói", flows: "Công việc & cứu hộ", library: "Tra cứu kho",
  update: "Cập nhật an toàn", transfer: "Chuyển máy", cleanup: "Dọn kho"
};
const $ = (selector, root = document) => root.querySelector(selector);
const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];
const api = async (url, options) => {
  const response = await fetch(url, options);
  let payload = {};
  try { payload = await response.json(); } catch (_) {}
  if (!response.ok) throw new Error(payload.error || "Không thể đọc dữ liệu");
  return payload;
};
const formatSize = bytes => {
  if (!Number.isFinite(bytes) || bytes <= 0) return "0 B";
  const units = ["B", "KB", "MB", "GB", "TB"];
  const level = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / 1024 ** level).toFixed(level > 1 ? 1 : 0)} ${units[level]}`;
};
const formatDate = stamp => new Intl.DateTimeFormat("vi-VN", { day: "2-digit", month: "2-digit", year: "2-digit" }).format(new Date(stamp * 1000));
const formatDuration = seconds => {
  const value = Math.max(0, Math.round(Number(seconds) || 0));
  return value < 60 ? `${value}s` : `${Math.floor(value / 60)}m ${value % 60}s`;
};
const escapeHtml = value => String(value).replace(/[&<>"']/g, char => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;" }[char]));
const normalizeSearch = value => String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
const toast = message => {
  const el = $("#toast"); el.textContent = message; el.classList.add("show");
  clearTimeout(window.toastTimer); window.toastTimer = setTimeout(() => el.classList.remove("show"), 2500);
};

async function restartStudio() {
  const button = $("#restart-studio");
  button.disabled = true;
  button.textContent = "Đang cập nhật…";
  try {
    await api("/api/restart", { method: "POST" });
  } catch (_) {}
  const started = Date.now();
  const waitForServer = async () => {
    try {
      const response = await fetch(`/api/health?t=${Date.now()}`, { cache: "no-store" });
      if (response.ok && Date.now() - started > 800) {
        window.location.replace(`/?updated=${Date.now()}`);
        return;
      }
    } catch (_) {}
    if (Date.now() - started < 15000) setTimeout(waitForServer, 400);
    else {
      button.disabled = false;
      button.textContent = "↻ Cập nhật Studio";
      toast("Chưa tự mở lại được. Hãy dùng file MỞ_HOTKEYVIP_STUDIO.bat");
    }
  };
  setTimeout(waitForServer, 700);
}

function navigate(page) {
  $$(".nav-item").forEach(btn => btn.classList.toggle("active", btn.dataset.page === page));
  $$(".page").forEach(el => el.classList.toggle("active", el.id === `page-${page}`));
  $("#page-title").textContent = titles[page];
  window.scrollTo({ top: 0, behavior: "smooth" });
  if (page === "voice") showVoiceControlPage();
}

async function openSubmitProfiles() {
  const button = $("#open-submit-profiles");
  button.disabled = true;
  try {
    await api("/api/submit-profiles/open", { method: "POST" });
    toast("Đã mở Quản lý Submit URL.");
  } catch (error) {
    toast(error.message);
  } finally {
    button.disabled = false;
  }
}

async function showVoiceControlPage() {
  try {
    const status = await refreshVoiceStatus(true);
    if (!status.running) {
      $("#voice-frame").hidden = true;
      $("#voice-placeholder").classList.remove("hidden");
      $("#voice-placeholder h3").textContent = "Voice Control đang tắt";
      $("#voice-placeholder p").textContent = "Bấm Khởi động khi muốn sử dụng.";
    }
  } catch (error) {
    toast(error.message);
  }
}

async function refreshVoiceStatus(loadFrame = false) {
  const status = await api("/api/voice/status");
  if (document.activeElement !== $("#voice-launcher-path")) $("#voice-launcher-path").value = status.launcherPath;
  if (document.activeElement !== $("#voice-interface-url")) $("#voice-interface-url").value = status.url;
  $("#voice-status").classList.toggle("running", status.running);
  $("#voice-status strong").textContent = status.running ? "Đang chạy" : "Chưa khởi động";
  if (status.running && loadFrame) {
    const frame = $("#voice-frame");
    if (frame.src !== status.url) frame.src = status.url;
    frame.hidden = false;
    $("#voice-placeholder").classList.add("hidden");
  }
  return status;
}

async function pickVoiceLauncher() {
  const button = $("#pick-voice-launcher");
  button.disabled = true;
  button.textContent = "Đang mở…";
  try {
    const result = await api("/api/voice/pick-launcher");
    if (result.path) $("#voice-launcher-path").value = result.path;
  } catch (error) {
    toast(error.message);
  } finally {
    button.disabled = false;
    button.textContent = "Chọn file chạy…";
  }
}

async function saveVoiceConfig() {
  try {
    const result = await api("/api/voice/config", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        launcherPath: $("#voice-launcher-path").value.trim(),
        url: $("#voice-interface-url").value.trim(),
      }),
    });
    toast(result.restartRequired
      ? "Đã lưu. Tắt toàn bộ rồi Khởi động lại để dùng file mới."
      : "Đã lưu cấu hình Voice Control.");
    await refreshVoiceStatus(false);
  } catch (error) {
    toast(error.message);
  }
}

async function startVoiceControl() {
  $("#voice-placeholder").classList.remove("hidden");
  $("#voice-frame").hidden = true;
  $("#voice-placeholder h3").textContent = "Đang khởi động Voice Control…";
  try {
    let status = await refreshVoiceStatus(true);
    if (!status.running) {
      await api("/api/voice/start", { method: "POST" });
      for (let attempt = 0; attempt < 20; attempt += 1) {
        await new Promise(resolve => setTimeout(resolve, 350));
        status = await refreshVoiceStatus(true);
        if (status.running) break;
      }
    }
    if (!status.running) throw new Error("Voice Control chưa mở được server ở cổng 8766");
  } catch (error) {
    $("#voice-placeholder h3").textContent = "Không khởi động được Voice Control";
    $("#voice-placeholder p").textContent = error.message;
    toast(error.message);
  }
}

async function stopVoiceControl() {
  try {
    await api("/api/voice/stop", { method: "POST" });
    $("#voice-frame").src = "about:blank";
    $("#voice-frame").hidden = true;
    $("#voice-placeholder").classList.remove("hidden");
    $("#voice-placeholder h3").textContent = "Voice Control đã tắt toàn bộ";
    $("#voice-placeholder p").textContent = "Bấm Khởi động khi muốn dùng lại.";
    await refreshVoiceStatus(false);
  } catch (error) {
    toast(error.message);
  }
}

async function setVoiceOverlay(action) {
  try {
    await api("/api/voice/overlay", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action }),
    });
    toast(action === "hide" ? "Đã ẩn bảng nổi Voice Control." : "Đã hiện bảng nổi Voice Control.");
  } catch (error) {
    toast(error.message);
  }
}

async function loadData() {
  try {
    const [summary, catalog, risks] = await Promise.all([api("/api/summary"), api("/api/catalog"), api("/api/risks")]);
    state.summary = summary; state.catalog = catalog;
    renderSummary(summary, risks); renderFlows(catalog); renderCategories(summary); renderCleanup(summary);
    await searchFiles();
    await previewPackage("core");
    await loadLaunchers();
    if (summary.scanning || !summary.scannedAt) window.setTimeout(loadData, 1800);
  } catch (error) {
    toast("Không đọc được kho. Hãy mở lại app.");
  }
}

async function loadLaunchers() {
  try {
    [state.launchers, state.launcherGroups] = await Promise.all([
      api("/api/launchers"),
      api("/api/launcher-groups"),
    ]);
    renderLaunchers();
  } catch (_error) {
    $("#launcher-grid").innerHTML = '<div class="launcher-loading">Không tải được danh sách nút.</div>';
  }
}

function renderLaunchers() {
  const items = state.launchers || [];
  const favoriteGroup = "★ Cần dùng";
  const preferredGroups = ["Tự động", "Cứu hộ thủ công", "Công cụ", "Sau khi đăng", "Chưa phân nhóm"];
  const grouped = new Map((state.launcherGroups || []).map(group => [group, []]));
  if (!grouped.size) grouped.set("Chưa phân nhóm", []);
  grouped.set(favoriteGroup, items.filter(item => item.favorite));
  items.forEach(item => {
    const group = item.group || "Chưa phân nhóm";
    if (!grouped.has(group)) grouped.set(group, []);
    grouped.get(group).push(item);
  });
  const groups = [favoriteGroup, ...[...grouped.keys()].filter(group => group !== favoriteGroup)];
  if (!groups.includes(state.activeLauncherGroup)) state.activeLauncherGroup = groups[0];
  $("#launcher-tabs").innerHTML = groups.map(group =>
    `<button class="launcher-tab ${group === state.activeLauncherGroup ? "active" : ""}" draggable="true" data-launcher-tab="${escapeHtml(group)}" title="Kéo trái/phải để đổi ưu tiên nhóm">${escapeHtml(group)}<small>${grouped.get(group).length}</small></button>`
  ).join("");
  $$("[data-launcher-tab]").forEach(tab => tab.addEventListener("click", () => {
    state.activeLauncherGroup = tab.dataset.launcherTab;
    renderLaunchers();
  }));
  $$("[data-launcher-tab]").forEach(tab => {
    tab.addEventListener("dragstart", event => {
      event.stopPropagation();
      tab.classList.add("group-dragging");
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("application/x-launcher-group", tab.dataset.launcherTab);
    });
    tab.addEventListener("dragend", () => {
      tab.classList.remove("group-dragging");
      $$(".launcher-tab").forEach(item => item.classList.remove("drag-target"));
    });
    tab.addEventListener("dragover", event => {
      event.preventDefault();
      tab.classList.add("drag-target");
    });
    tab.addEventListener("dragleave", () => tab.classList.remove("drag-target"));
    tab.addEventListener("drop", async event => {
      event.preventDefault();
      tab.classList.remove("drag-target");
      const sourceGroup = event.dataTransfer.getData("application/x-launcher-group");
      if (sourceGroup) {
        await reorderLauncherGroup(sourceGroup, tab.dataset.launcherTab);
        return;
      }
      if (event.dataTransfer.getData("application/x-favorite-order")) return;
      const id = event.dataTransfer.getData("text/plain");
      if (id) await moveLauncherToGroup(id, tab.dataset.launcherTab);
    });
  });
  const searchQuery = normalizeSearch(state.launcherQuery.trim());
  const viewingSearch = Boolean(searchQuery);
  const viewingFavorites = !viewingSearch && state.activeLauncherGroup === favoriteGroup;
  const rows = (viewingSearch ? items.filter(item => normalizeSearch([
    item.name, item.description, item.group, item.path, String(item.path || "").split(/[\\/]/).pop()
  ].join(" ")).includes(searchQuery)) : grouped.get(state.activeLauncherGroup)).sort((a, b) =>
    Number(viewingFavorites ? (a.favoriteOrder || 9999) : (a.order || 9999)) -
    Number(viewingFavorites ? (b.favoriteOrder || 9999) : (b.order || 9999))
  );
  $("#launcher-search-count").textContent = viewingSearch ? `${rows.length} kết quả` : "";
  const realGroups = groups.filter(group => group !== favoriteGroup);
  const groupOptions = realGroups.map(group => `<option value="${escapeHtml(group)}">${escapeHtml(group)}</option>`).join("");
  const statusLabel = item => {
    if (!item.exists || !item.validType) return ["missing", "Thiếu file"];
    if (item.runningPids?.length) return ["running", `Đang chạy ${formatDuration(Date.now() / 1000 - item.lastLogStartedAt)}`];
    if (item.lastLogState === "success") return ["success", `Thành công · ${formatDuration(item.lastLogDuration)}`];
    if (item.lastLogState === "error") return ["error", `Bị lỗi · ${formatDuration(item.lastLogDuration)}`];
    if (item.lastLogState === "stopped") return ["stopped", `Đã dừng · ${formatDuration(item.lastLogDuration)}`];
    return ["ready", "Sẵn sàng"];
  };
  $("#launcher-grid").innerHTML = `<section class="launcher-group">
      <div class="launcher-group-body" data-launcher-group="${escapeHtml(viewingSearch ? "__search__" : state.activeLauncherGroup)}">${rows.length ? rows.map((item, index) => `
        <div class="launcher-row" draggable="${viewingSearch ? "false" : "true"}" data-launcher-row="${escapeHtml(item.id)}">
          <input class="priority-input" type="number" min="1" max="${rows.length}" value="${index + 1}"
            data-priority-launcher="${escapeHtml(item.id)}" title="Nhập vị trí mới rồi nhấn Enter" ${viewingSearch ? "disabled" : ""}>
          <div class="launcher-name"><strong>${escapeHtml(item.name)}</strong><small>${escapeHtml(item.description || "Không có mô tả")}</small></div>
          <label class="favorite-toggle" title="Hiện trong nhóm Cần dùng"><input type="checkbox" data-favorite-launcher="${escapeHtml(item.id)}" ${item.favorite ? "checked" : ""}><span>★</span></label>
          <span class="console-badge">${item.showConsole ? "Hiện CMD" : "Chạy ẩn"}</span>
          ${(() => { const [kind, label] = statusLabel(item); return `<span class="path-state ${kind}">${label}</span>`; })()}
          <select class="quick-group-select" data-move-launcher="${escapeHtml(item.id)}" title="Chuyển nhanh sang nhóm khác">${groupOptions}</select>
          ${item.runningPids?.length
            ? `<button class="run-button stop-button" data-stop-launcher="${escapeHtml(item.id)}">Dừng</button>`
            : `<button class="run-button" data-run-launcher="${escapeHtml(item.id)}" ${item.exists && item.validType ? "" : "disabled"}>${escapeHtml(item.actionLabel || "Mở")}</button>`}
          ${/\.pyw?$/i.test(item.path || "") ? `<button class="log-button state-${escapeHtml(item.lastLogState || "none")}" data-history-launcher="${escapeHtml(item.id)}" ${item.hasLog ? "" : "disabled"} title="Xem 10 lượt chạy gần nhất">Lịch sử</button>` : `<span class="log-spacer"></span>`}
          <button class="launcher-edit" data-edit-launcher="${escapeHtml(item.id)}" title="Sửa">⋯</button>
        </div>`).join("") : `<div class="launcher-search-empty">Không tìm thấy nút phù hợp.</div>`}</div>
    </section>`;
  $$("[data-move-launcher]").forEach(select => {
    const item = items.find(row => row.id === select.dataset.moveLauncher);
    select.value = item?.group || "Chưa phân nhóm";
    select.addEventListener("change", () => moveLauncherToGroup(select.dataset.moveLauncher, select.value));
  });
  $$("[data-favorite-launcher]").forEach(input => input.addEventListener("change", () =>
    toggleLauncherFavorite(input.dataset.favoriteLauncher, input.checked)
  ));
  $$("[data-priority-launcher]").forEach(input => {
    const row = input.closest("[data-launcher-row]");
    input.addEventListener("focus", () => row.setAttribute("draggable", "false"));
    input.addEventListener("blur", () => {
      row.setAttribute("draggable", "true");
      setLauncherPosition(input.dataset.priorityLauncher, input.value);
    });
    input.addEventListener("keydown", event => {
      if (event.key === "Enter") input.blur();
    });
  });
  $$("[data-edit-launcher]").forEach(button => button.addEventListener("click", () => openLauncherModal(button.dataset.editLauncher)));
  $$("[data-run-launcher]").forEach(button => button.addEventListener("click", () => runLauncher(button.dataset.runLauncher, button)));
  $$("[data-stop-launcher]").forEach(button => button.addEventListener("click", () => stopLauncher(button.dataset.stopLauncher)));
  $$("[data-history-launcher]").forEach(button => button.addEventListener("click", () => showLauncherHistory(button.dataset.historyLauncher)));
  if (!viewingSearch) bindLauncherDrag();
}

async function setLauncherPosition(id, requestedPosition) {
  const body = $(".launcher-group-body");
  if (!body || body.dataset.launcherGroup === "__search__") return;
  const rows = $$("[data-launcher-row]", body);
  const currentIndex = rows.findIndex(row => row.dataset.launcherRow === id);
  if (currentIndex < 0) return;
  const targetIndex = Math.max(0, Math.min(rows.length - 1, Number(requestedPosition || 1) - 1));
  if (currentIndex === targetIndex) {
    renderLaunchers();
    return;
  }
  const orderedIds = rows.map(row => row.dataset.launcherRow);
  orderedIds.splice(currentIndex, 1);
  orderedIds.splice(targetIndex, 0, id);
  const items = orderedIds.map((launcherId, index) => ({ id: launcherId, order: (index + 1) * 10 }));
  try {
    if (body.dataset.launcherGroup === "★ Cần dùng") {
      await api("/api/launchers/favorite-reorder", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ items }),
      });
    } else {
      await api("/api/launchers/reorder", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ items: items.map(item => ({ ...item, group: body.dataset.launcherGroup })) }),
      });
    }
    await loadLaunchers();
    toast(`Đã chuyển tới vị trí ${targetIndex + 1}.`);
  } catch (_error) {
    await loadLaunchers();
    toast("Không đổi được vị trí.");
  }
}

async function toggleLauncherFavorite(id, favorite) {
  try {
    await api("/api/launchers/favorite", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ id, favorite }),
    });
    await loadLaunchers();
    toast(favorite ? "Đã thêm vào Cần dùng." : "Đã bỏ khỏi Cần dùng.");
  } catch (_error) {
    await loadLaunchers();
    toast("Không lưu được đánh dấu.");
  }
}

async function reorderLauncherGroup(source, target) {
  if (!source || !target || source === target) return;
  const groups = [...(state.launcherGroups || [])];
  const from = groups.indexOf(source);
  const to = groups.indexOf(target);
  if (from < 0 || to < 0) return;
  groups.splice(from, 1);
  groups.splice(to, 0, source);
  try {
    await api("/api/launcher-groups/reorder", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ groups }),
    });
    state.launcherGroups = groups;
    renderLaunchers();
    toast("Đã lưu thứ tự ưu tiên nhóm.");
  } catch (_error) {
    await loadLaunchers();
    toast("Không lưu được thứ tự nhóm.");
  }
}

async function moveLauncherToGroup(id, group) {
  try {
    await api("/api/launchers/reorder", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ items: [{ id, group, order: 9999 }] }),
    });
    state.activeLauncherGroup = group;
    await loadLaunchers();
    toast("Đã chuyển nút sang nhóm mới.");
  } catch (_error) {
    await loadLaunchers();
    toast("Không chuyển được nhóm.");
  }
}

function bindLauncherDrag() {
  let dragged = null;
  $$("[data-launcher-row]").forEach(row => {
    row.addEventListener("dragstart", event => {
      dragged = row; row.classList.add("dragging");
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", row.dataset.launcherRow);
      if (state.activeLauncherGroup === "★ Cần dùng") {
        event.dataTransfer.setData("application/x-favorite-order", "true");
      }
    });
    row.addEventListener("dragend", () => {
      row.classList.remove("dragging");
      $$(".launcher-group-body").forEach(body => body.classList.remove("drag-over"));
      dragged = null;
    });
  });
  $$(".launcher-group-body").forEach(body => {
    body.addEventListener("dragover", event => {
      event.preventDefault();
      if (!dragged) return;
      body.classList.add("drag-over");
      const candidates = $$("[data-launcher-row]:not(.dragging)", body);
      const next = candidates.find(row => event.clientY < row.getBoundingClientRect().top + row.offsetHeight / 2);
      body.insertBefore(dragged, next || null);
    });
    body.addEventListener("dragleave", event => {
      if (!body.contains(event.relatedTarget)) body.classList.remove("drag-over");
    });
    body.addEventListener("drop", async event => {
      event.preventDefault(); body.classList.remove("drag-over");
      await persistLauncherOrder();
    });
  });
}

async function persistLauncherOrder() {
  const items = [];
  const favoriteBody = $(".launcher-group-body[data-launcher-group='★ Cần dùng']");
  if (favoriteBody) {
    $$("[data-launcher-row]", favoriteBody).forEach((row, index) => items.push({
      id: row.dataset.launcherRow,
      order: (index + 1) * 10,
    }));
    try {
      await api("/api/launchers/favorite-reorder", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ items }),
      });
      await loadLaunchers();
      toast("Đã lưu thứ tự trong Cần dùng.");
    } catch (_error) {
      await loadLaunchers();
      toast("Không lưu được thứ tự Cần dùng.");
    }
    return;
  }
  $$(".launcher-group-body").forEach(body => {
    $$("[data-launcher-row]", body).forEach((row, index) => items.push({
      id: row.dataset.launcherRow,
      group: body.dataset.launcherGroup,
      order: (index + 1) * 10,
    }));
  });
  try {
    await api("/api/launchers/reorder", { method:"POST", headers:{ "Content-Type":"application/json" }, body:JSON.stringify({ items }) });
    await loadLaunchers(); toast("Đã lưu thứ tự mới.");
  } catch (_error) {
    await loadLaunchers(); toast("Không lưu được thứ tự.");
  }
}

async function renameLauncherGroup() {
  const oldName = state.activeLauncherGroup;
  if (!oldName) return;
  if (oldName === "★ Cần dùng") return toast("Đây là nhóm đánh dấu tự động.");
  const newName = prompt("Tên mới của nhóm:", oldName);
  if (!newName || newName.trim() === oldName) return;
  try {
    await api("/api/launchers/rename-group", {
      method:"POST",
      headers:{ "Content-Type":"application/json" },
      body:JSON.stringify({ oldName, newName:newName.trim() }),
    });
    state.activeLauncherGroup = newName.trim();
    await loadLaunchers();
    toast("Đã đổi tên nhóm.");
  } catch (_error) {
    toast("Không đổi được tên nhóm.");
  }
}

async function deleteLauncherGroup() {
  const name = state.activeLauncherGroup;
  if (!name || name === "Chưa phân nhóm" || name === "★ Cần dùng") {
    toast("Không thể xóa nhóm hệ thống này.");
    return;
  }
  if (!confirm(`Xóa nhóm “${name}”? Các nút trong nhóm sẽ được chuyển về “Chưa phân nhóm”.`)) return;
  try {
    await api("/api/launcher-groups/delete", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name }),
    });
    state.activeLauncherGroup = "Chưa phân nhóm";
    await loadLaunchers();
    toast("Đã xóa nhóm. Các nút vẫn được giữ nguyên.");
  } catch (_error) {
    toast("Không xóa được nhóm.");
  }
}

async function addLauncherGroup() {
  const name = prompt("Tên nhóm mới:");
  if (!name || !name.trim()) return;
  try {
    await api("/api/launcher-groups/add", {
      method:"POST",
      headers:{ "Content-Type":"application/json" },
      body:JSON.stringify({ name:name.trim() }),
    });
    state.activeLauncherGroup = name.trim();
    await loadLaunchers();
    toast("Đã thêm nhóm mới.");
  } catch (_error) {
    toast("Không thêm được nhóm; có thể tên đã tồn tại.");
  }
}

async function pickLauncherFile() {
  const button = $("#pick-launcher-file");
  button.disabled = true;
  button.textContent = "Đang mở…";
  try {
    const result = await api("/api/pick-python-file");
    if (result.path) $("#launcher-path").value = result.path;
  } catch (_error) {
    toast("Không mở được hộp chọn file.");
  } finally {
    button.disabled = false;
    button.textContent = "Chọn file…";
  }
}

async function openLauncherFolder() {
  const path = $("#launcher-path").value.trim();
  if (!path) return toast("Chưa có đường dẫn file.");
  try {
    await api("/api/open-containing-folder", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path }),
    });
  } catch (_error) {
    toast("Không mở được thư mục; hãy kiểm tra lại đường dẫn.");
  }
}

function openLauncherModal(id = "") {
  const item = (state.launchers || []).find(row => row.id === id);
  const groups = state.launcherGroups || ["Chưa phân nhóm"];
  $("#launcher-group").innerHTML = groups.map(group =>
    `<option value="${escapeHtml(group)}">${escapeHtml(group)}</option>`
  ).join("");
  $("#launcher-form").reset();
  $("#launcher-id").value = item?.id || "";
  $("#launcher-name").value = item?.name || "";
  $("#launcher-path").value = item?.path || "";
  $("#launcher-description").value = item?.description || "";
  $("#launcher-group").value = item?.group || state.activeLauncherGroup || "Chưa phân nhóm";
  $("#launcher-console").checked = item ? item.showConsole !== false : true;
  $("#launcher-form-title").textContent = item ? "Sửa nút chạy" : "Thêm nút mới";
  $("#delete-launcher").classList.toggle("hidden", !item);
  $("#launcher-error").textContent = "";
  $("#launcher-modal").classList.remove("hidden");
  setTimeout(() => $("#launcher-name").focus(), 80);
}

function closeLauncherModal() { $("#launcher-modal").classList.add("hidden"); }

async function saveLauncher(event) {
  event.preventDefault();
  $("#launcher-error").textContent = "";
  const payload = {
    id: $("#launcher-id").value || undefined,
    name: $("#launcher-name").value.trim(),
    path: $("#launcher-path").value.trim(),
    description: $("#launcher-description").value.trim(),
    group: $("#launcher-group").value.trim() || "Chưa phân nhóm",
    order: (state.launchers || []).find(row => row.id === $("#launcher-id").value)?.order || 9999,
    showConsole: $("#launcher-console").checked,
  };
  try {
    await api("/api/launchers/save", { method:"POST", headers:{ "Content-Type":"application/json" }, body:JSON.stringify(payload) });
    closeLauncherModal(); await loadLaunchers(); toast("Đã lưu nút chạy.");
  } catch (error) {
    $("#launcher-error").textContent = error.message || "Không lưu được nút.";
  }
}

async function deleteLauncher() {
  const id = $("#launcher-id").value;
  if (!id || !confirm("Xóa nút này khỏi GUI? File Python sẽ không bị xóa.")) return;
  await api("/api/launchers/delete", { method:"POST", headers:{ "Content-Type":"application/json" }, body:JSON.stringify({ id }) });
  closeLauncherModal(); await loadLaunchers(); toast("Đã xóa nút khỏi GUI. File Python vẫn nguyên vẹn.");
}

async function runLauncher(id, button) {
  const originalLabel = button.textContent;
  button.disabled = true; button.textContent = "Đang mở…";
  try {
    const result = await api("/api/launchers/run", { method:"POST", headers:{ "Content-Type":"application/json" }, body:JSON.stringify({ id }) });
    toast(result.logged
      ? `Đã chạy “${result.name}” và bắt đầu ghi log.`
      : `Đã mở “${result.name}”.`);
    if (result.logged) setTimeout(loadLaunchers, 700);
  } catch (error) {
    toast(error.message || "Không chạy được file.");
  } finally {
    button.disabled = false; button.textContent = originalLabel;
  }
}

async function stopLauncher(id) {
  try {
    await api("/api/launchers/stop", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ id }),
    });
    toast("Đã dừng tiến trình.");
    await loadLaunchers();
  } catch (error) {
    toast(error.message || "Không dừng được tiến trình.");
  }
}

async function showLauncherHistory(id) {
  try {
    const rows = await api(`/api/launchers/history?id=${encodeURIComponent(id)}`);
    if (!rows.length) return toast("Nút này chưa có lịch sử.");
    const stateNames = { running: "ĐANG CHẠY", success: "THÀNH CÔNG", error: "BỊ LỖI", stopped: "ĐÃ DỪNG" };
    const launcher = (state.launchers || []).find(item => item.id === id);
    state.historyLauncherId = id;
    state.historyLogName = "";
    $("#history-title").textContent = launcher ? `Lịch sử · ${launcher.name}` : "Lịch sử chạy";
    $("#history-runs").innerHTML = rows.map((row, index) => {
      const when = new Date(Number(row.startedAt || 0) * 1000).toLocaleString("vi-VN");
      const stateClass = ["running", "success", "error", "stopped"].includes(row.state) ? row.state : "unknown";
      const duration = row.state === "running"
        ? Date.now() / 1000 - Number(row.startedAt || Date.now() / 1000)
        : row.duration;
      return `<button type="button" class="history-run state-${stateClass}" data-history-run="${escapeHtml(row.name)}">
        <span class="history-run-index">${index + 1}</span>
        <span><strong>${escapeHtml(stateNames[row.state] || row.state || "KHÔNG RÕ")}</strong><small>${escapeHtml(when)}</small></span>
        <span class="history-run-duration">${escapeHtml(formatDuration(duration))}</span>
      </button>`;
    }).join("");
    $("#history-modal").classList.remove("hidden");
    $$("[data-history-run]", $("#history-runs")).forEach(button =>
      button.addEventListener("click", () => selectLauncherHistory(button.dataset.historyRun))
    );
    await selectLauncherHistory(rows[0].name);
  } catch (error) {
    toast(error.message || "Không đọc được lịch sử.");
  }
}

function closeLauncherHistory() {
  $("#history-modal").classList.add("hidden");
  state.historyLauncherId = "";
  state.historyLogName = "";
}

async function selectLauncherHistory(name) {
  if (!state.historyLauncherId || !name) return;
  state.historyLogName = name;
  $$("[data-history-run]", $("#history-runs")).forEach(button =>
    button.classList.toggle("active", button.dataset.historyRun === name)
  );
  $("#history-log-title").textContent = name;
  $("#history-log-meta").textContent = "Đang đọc nội dung…";
  $("#history-log-content").textContent = "Đang tải log…";
  try {
    const log = await api(`/api/launchers/log?id=${encodeURIComponent(state.historyLauncherId)}&name=${encodeURIComponent(name)}`);
    const lineCount = log.content ? log.content.split(/\r?\n/).length : 0;
    $("#history-log-title").textContent = log.name;
    $("#history-log-meta").textContent = `${formatSize(log.size)} · ${lineCount.toLocaleString("vi-VN")} dòng · cập nhật ${new Date(log.modifiedAt * 1000).toLocaleString("vi-VN")}`;
    $("#history-log-content").textContent = log.content || "(Log trống)";
    $("#history-log-content").scrollTop = 0;
  } catch (error) {
    $("#history-log-meta").textContent = "Không đọc được lượt chạy này";
    $("#history-log-content").textContent = error.message || "Không đọc được log.";
  }
}

async function copyLauncherLog() {
  const content = $("#history-log-content").textContent;
  if (!content || !state.historyLogName) return;
  try {
    await navigator.clipboard.writeText(content);
    toast("Đã sao chép nội dung log.");
  } catch (_error) {
    const fallback = document.createElement("textarea");
    fallback.value = content;
    fallback.style.position = "fixed";
    fallback.style.opacity = "0";
    document.body.appendChild(fallback);
    fallback.select();
    const copied = document.execCommand("copy");
    fallback.remove();
    toast(copied ? "Đã sao chép nội dung log." : "Không sao chép được. Hãy bôi đen log và nhấn Ctrl+C.");
  }
}

async function openLauncherLog(id = state.historyLauncherId, name = state.historyLogName) {
  try {
    await api("/api/launchers/open-log", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ id, name }),
    });
  } catch (error) {
    toast(error.message || "Không mở được log.");
  }
}

function renderSummary(data, risks) {
  $("#metric-size").textContent = formatSize(data.size);
  $("#metric-files").textContent = `${data.files.toLocaleString("vi-VN")} file đã lập chỉ mục`;
  const review = data.categories["Cần xem xét"] || { files: 0 };
  $("#metric-review").textContent = review.files.toLocaleString("vi-VN");
  $("#hero-copy").textContent = data.scanning && !data.scannedAt
    ? `App đã mở và đang lập bản đồ kho ${data.root} ở chế độ chỉ đọc…`
    : `Đã đọc ${data.files.toLocaleString("vi-VN")} file trong ${data.duration.toFixed(1)} giây. App chỉ quan sát kho tại ${data.root}.`;
  const max = Math.max(...data.folders.map(row => row.size), 1);
  $("#folder-bars").innerHTML = data.folders.slice(0, 7).map(row => `
    <div class="bar-row"><strong title="${escapeHtml(row.name)}">${escapeHtml(row.name)}</strong>
      <div class="bar-track"><div class="bar-fill" style="width:${Math.max(row.size / max * 100, 1)}%"></div></div>
      <span>${formatSize(row.size)}</span></div>`).join("");
  $("#risk-list").innerHTML = risks.slice(0, 4).map(risk => `
    <div class="risk ${risk.level}"><span class="risk-dot"></span><div><strong>${escapeHtml(risk.title)}</strong>
    <p>${escapeHtml(risk.detail)}</p></div><em>${risk.level === "high" ? "Ưu tiên" : "Cần xem"}</em></div>`).join("");
}

function renderFlows(items) {
  const icons = { control:"⌘", "prepare-writing":"01", writing:"✦", "prepare-publishing":"02", publishing:"⇧", workbook:"▦" };
  $("#flow-grid").innerHTML = items.map(item => {
    const allExist = item.primary.exists && item.tools.every(tool => tool.exists);
    return `<article class="flow-card" data-flow="${item.id}">
      <div class="flow-icon">${icons[item.id] || "•"}</div><h3>${escapeHtml(item.name)}</h3><p>${escapeHtml(item.description)}</p>
      <div class="flow-meta"><span>${item.steps.length} bước · ${item.tools.length} công cụ cứu hộ</span>
      <span class="${allExist ? "ok" : ""}">${allExist ? "✓ Đã nhận diện" : "△ Cần xác nhận"}</span></div></article>`;
  }).join("");
  $$(".flow-card").forEach(card => card.addEventListener("click", () => openFlow(card.dataset.flow)));
}

function openFlow(id) {
  const item = state.catalog.find(flow => flow.id === id);
  if (!item) return;
  const drawer = $("#flow-detail");
  drawer.classList.remove("hidden");
  drawer.innerHTML = `<div class="drawer-head"><div><h3>${escapeHtml(item.name)}</h3><p>Điểm chạy được nhận diện: ${escapeHtml(item.primary.path)}</p></div><button id="close-drawer">Đóng ×</button></div>
    <div class="step-line">${item.steps.map((step, i) => `<div class="step" data-n="${i+1}">${escapeHtml(step)}</div>`).join("")}</div>
    <div class="tool-list">${item.tools.length ? item.tools.map(tool => `<span class="tool-pill">${tool.exists ? "✓" : "△"} ${escapeHtml(tool.path.split("\\").pop())}</span>`).join("") : '<span class="tool-pill">Chưa khai báo công cụ cứu hộ riêng</span>'}</div>`;
  $("#close-drawer").onclick = () => drawer.classList.add("hidden");
  drawer.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function renderCategories(summary) {
  const categories = ["Tất cả", ...Object.keys(summary.categories)];
  $("#category-chips").innerHTML = categories.map((name, i) => `<button class="chip ${i === 0 ? "active" : ""}" data-category="${escapeHtml(name)}">${escapeHtml(name)}</button>`).join("");
  $$(".chip").forEach(chip => chip.addEventListener("click", () => {
    $$(".chip").forEach(item => item.classList.remove("active")); chip.classList.add("active");
    state.category = chip.dataset.category; searchFiles();
  }));
}

async function searchFiles() {
  const rows = await api(`/api/files?q=${encodeURIComponent(state.query)}&category=${encodeURIComponent(state.category)}&limit=100`);
  $("#result-caption").textContent = `${rows.length} kết quả đầu tiên · không có file nào bị thay đổi`;
  $("#file-results").innerHTML = rows.length ? rows.map(item => `
    <div class="file-row"><div class="file-name"><strong title="${escapeHtml(item.path)}">${escapeHtml(item.name)}</strong><small>${escapeHtml(item.folder)} · ${escapeHtml(item.extension)}</small></div>
    <span class="category-label" title="${escapeHtml(item.reason)}">${escapeHtml(item.category)}</span>
    <span>${formatSize(item.size)}</span><span>${formatDate(item.modified)}</span></div>`).join("") : `<div class="file-empty">Không tìm thấy file phù hợp.</div>`;
}

async function analyzeUpdate(file) {
  const result = await api("/api/update-preview", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: file.name, size: file.size })
  });
  $("#update-preview").innerHTML = `<div class="preview-file"><div class="file-badge">${escapeHtml((file.name.split(".").pop() || "FILE").toUpperCase())}</div>
    <div><strong>${escapeHtml(file.name)}</strong><small>${formatSize(file.size)} · ${escapeHtml(result.mode)}</small></div></div>
    <div class="target-box"><span>Chức năng được đề xuất</span><strong>${escapeHtml(result.suggestedTarget)}</strong></div>
    <div class="check-list">${result.checks.map(check => `<div class="check-item ${check.state === "warn" ? "warn" : ""}"><i>${check.state === "pass" ? "✓" : "!"}</i>${escapeHtml(check.label)}</div>`).join("")}</div>`;
}

async function previewPackage(mode) {
  state.packageMode = mode;
  const result = await api("/api/package-preview", { method:"POST", headers:{ "Content-Type":"application/json" }, body:JSON.stringify({ mode }) });
  $("#package-size").textContent = formatSize(result.includeSize);
  $("#excluded-size").textContent = formatSize(result.excludedSize);
  $("#package-items").innerHTML = result.items.map(item => `<div class="package-item ${item.included ? "" : "excluded"}"><span>${item.included ? "✓" : "×"}</span>${escapeHtml(item.name)}</div>`).join("");
}

function renderCleanup(summary) {
  const temp = summary.categories["Tạm & có thể tạo lại"] || { files:0, size:0 };
  const history = summary.categories["Lịch sử / bản cũ"] || { files:0, size:0 };
  const tests = summary.categories["Thử nghiệm"] || { files:0, size:0 };
  const cards = [
    ["◌", temp.size, "Cache, log và file tạm", `${temp.files.toLocaleString("vi-VN")} file có dấu hiệu tạo lại được.`, "Cần tách dữ liệu đăng nhập trước"],
    ["↶", history.size, "Bản sao và lịch sử", `${history.files.toLocaleString("vi-VN")} file có tên backup, copy hoặc bản cũ.`, "Nên đóng gói thay vì xóa"],
    ["◇", tests.size, "Kết quả thử nghiệm", `${tests.files.toLocaleString("vi-VN")} file nằm trong vùng test hoặc chẩn đoán.`, "Kiểm tra ngày sử dụng gần nhất"],
  ];
  $("#cleanup-cards").innerHTML = cards.map(row => `<article class="cleanup-card"><div class="clean-icon">${row[0]}</div><strong>${formatSize(row[1])}</strong><h3>${row[2]}</h3><p>${row[3]}</p><footer>${row[4]}</footer></article>`).join("");
}

function bindEvents() {
  $$(".nav-item[data-page]").forEach(btn => btn.addEventListener("click", () => navigate(btn.dataset.page)));
  $("#open-submit-profiles").addEventListener("click", openSubmitProfiles);
  $$("[data-goto]").forEach(btn => btn.addEventListener("click", () => navigate(btn.dataset.goto)));
  $("#restart-studio").addEventListener("click", restartStudio);
  $("#rescan").addEventListener("click", async () => {
    $("#rescan").textContent = "Đang quét…";
    const summary = await api("/api/rescan");
    const risks = await api("/api/risks");
    state.summary = summary; renderSummary(summary, risks); renderCleanup(summary); renderCategories(summary); await searchFiles();
    $("#rescan").textContent = "↻ Quét lại"; toast("Đã cập nhật bản đồ kho. Không có file nào bị thay đổi.");
  });
  const search = $("#file-search");
  search.addEventListener("input", event => { state.query = event.target.value; clearTimeout(window.searchTimer); window.searchTimer = setTimeout(searchFiles, 220); });
  const picker = $("#file-picker");
  $("#choose-file").addEventListener("click", () => picker.click());
  picker.addEventListener("change", () => picker.files[0] && analyzeUpdate(picker.files[0]));
  const drop = $("#drop-zone");
  ["dragenter","dragover"].forEach(type => drop.addEventListener(type, event => { event.preventDefault(); drop.classList.add("drag"); }));
  ["dragleave","drop"].forEach(type => drop.addEventListener(type, event => { event.preventDefault(); drop.classList.remove("drag"); }));
  drop.addEventListener("drop", event => event.dataTransfer.files[0] && analyzeUpdate(event.dataTransfer.files[0]));
  $$(".transfer-option").forEach(option => option.addEventListener("click", () => {
    $$(".transfer-option").forEach(item => item.classList.remove("active")); option.classList.add("active"); previewPackage(option.dataset.mode);
  }));
  $$(".primary-button.disabled").forEach(btn => btn.addEventListener("click", () => toast("Bản thử nghiệm đang khóa thao tác ghi file.")));
  $("#add-launcher").addEventListener("click", () => openLauncherModal());
  $("#close-launcher-modal").addEventListener("click", closeLauncherModal);
  $("#cancel-launcher").addEventListener("click", closeLauncherModal);
  $("#launcher-modal").addEventListener("click", event => { if (event.target.id === "launcher-modal") closeLauncherModal(); });
  $("#close-history-modal").addEventListener("click", closeLauncherHistory);
  $("#history-modal").addEventListener("click", event => { if (event.target.id === "history-modal") closeLauncherHistory(); });
  $("#reload-history-log").addEventListener("click", () => selectLauncherHistory(state.historyLogName));
  $("#copy-history-log").addEventListener("click", copyLauncherLog);
  $("#open-history-log").addEventListener("click", () => openLauncherLog());
  $("#launcher-form").addEventListener("submit", saveLauncher);
  $("#delete-launcher").addEventListener("click", deleteLauncher);
  $("#add-launcher-group").addEventListener("click", addLauncherGroup);
  $("#rename-launcher-group").addEventListener("click", renameLauncherGroup);
  $("#delete-launcher-group").addEventListener("click", deleteLauncherGroup);
  $("#pick-launcher-file").addEventListener("click", pickLauncherFile);
  $("#open-launcher-folder").addEventListener("click", openLauncherFolder);
  $("#start-voice").addEventListener("click", startVoiceControl);
  $("#stop-voice").addEventListener("click", stopVoiceControl);
  $("#hide-voice-overlay").addEventListener("click", () => setVoiceOverlay("hide"));
  $("#show-voice-overlay").addEventListener("click", () => setVoiceOverlay("show"));
  $("#pick-voice-launcher").addEventListener("click", pickVoiceLauncher);
  $("#save-voice-config").addEventListener("click", saveVoiceConfig);
  $("#launcher-search").addEventListener("input", event => {
    state.launcherQuery = event.target.value;
    renderLaunchers();
  });
  document.addEventListener("keydown", event => {
    if (event.key === "Escape" && !$("#history-modal").classList.contains("hidden")) {
      closeLauncherHistory();
      return;
    }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "k") {
      event.preventDefault();
      $("#launcher-search").focus();
      $("#launcher-search").select();
    }
    if (event.key === "Escape" && document.activeElement === $("#launcher-search")) {
      $("#launcher-search").value = "";
      state.launcherQuery = "";
      renderLaunchers();
      $("#launcher-search").blur();
    }
  });
}

bindEvents();
loadLaunchers();
setInterval(() => {
  if ($("#page-launcher")?.classList.contains("active") && !document.hidden) loadLaunchers();
}, 2000);
