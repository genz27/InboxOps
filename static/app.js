const AUTO_REFRESH_MS = 5000;
const DEFAULT_MAILBOX_PAGE_SIZE = 8;
const WORKSPACE_STORAGE_KEY = "mail_console_workspace_v1";

const state = {
    authenticated: false,
    username: "",
    mailboxes: [],
    mailboxMeta: createDefaultMailboxMeta(),
    mailboxSearchTerm: "",
    mailboxPage: 1,
    mailboxPageSize: DEFAULT_MAILBOX_PAGE_SIZE,
    mailboxFetchToken: 0,
    mailboxSearchTimer: null,
    mailboxConnectionResults: {},
    selectedMailboxIds: new Set(),
    bulkPreferredMethod: "",
    selectedMailboxId: null,
    selectedMailbox: null,
    selectedMailboxDetail: null,
    activeMethod: null,
    folders: [],
    selectedFolder: "INBOX",
    overviewMethods: [],
    messages: [],
    messageMeta: createDefaultMessageMeta(),
    messagePage: 1,
    messagePageSize: 20,
    readState: "all",
    hasAttachmentsOnly: false,
    flaggedOnly: false,
    importance: "all",
    sortOrder: "desc",
    selectedMessageIds: new Set(),
    selectedMessageId: null,
    selectedMessage: null,
    searchTerm: "",
    autoRefreshTimer: null,
    searchTimer: null,
    isLoadingMessages: false,
    editorMode: "create",
    messageFetchToken: 0,
    detailFetchToken: 0,
    detailView: "text",
};

const elements = {
    authView: document.getElementById("authView"),
    dashboardView: document.getElementById("dashboardView"),
    sessionLabel: document.getElementById("sessionLabel"),
    currentMailboxLabel: document.getElementById("currentMailboxLabel"),
    currentMethodLabel: document.getElementById("currentMethodLabel"),
    autoRefreshLabel: document.getElementById("autoRefreshLabel"),
    lastSyncLabel: document.getElementById("lastSyncLabel"),
    logoutButton: document.getElementById("logoutButton"),
    statusText: document.getElementById("statusText"),
    loginForm: document.getElementById("loginForm"),
    loginButton: document.getElementById("loginButton"),
    mailboxList: document.getElementById("mailboxList"),
    mailboxSearchInput: document.getElementById("mailboxSearchInput"),
    mailboxResultInfo: document.getElementById("mailboxResultInfo"),
    mailboxPageInfo: document.getElementById("mailboxPageInfo"),
    mailboxPrevPageButton: document.getElementById("mailboxPrevPageButton"),
    mailboxNextPageButton: document.getElementById("mailboxNextPageButton"),
    sidebarFillState: document.getElementById("sidebarFillState"),
    selectedMailboxCount: document.getElementById("selectedMailboxCount"),
    toggleVisibleMailboxesButton: document.getElementById("toggleVisibleMailboxesButton"),
    clearSelectedMailboxesButton: document.getElementById("clearSelectedMailboxesButton"),
    testSelectedMailboxesButton: document.getElementById("testSelectedMailboxesButton"),
    bulkPreferredMethodSelect: document.getElementById("bulkPreferredMethodSelect"),
    applyPreferredMethodButton: document.getElementById("applyPreferredMethodButton"),
    deleteSelectedMailboxesButton: document.getElementById("deleteSelectedMailboxesButton"),
    testVisibleMailboxesButton: document.getElementById("testVisibleMailboxesButton"),
    mailboxItemTemplate: document.getElementById("mailboxItemTemplate"),
    newMailboxButton: document.getElementById("newMailboxButton"),
    importMailboxesButton: document.getElementById("importMailboxesButton"),
    editorDrawer: document.getElementById("editorDrawer"),
    editorModeLabel: document.getElementById("editorModeLabel"),
    closeEditorButton: document.getElementById("closeEditorButton"),
    mailboxForm: document.getElementById("mailboxForm"),
    saveMailboxButton: document.getElementById("saveMailboxButton"),
    resetMailboxButton: document.getElementById("resetMailboxButton"),
    testMailboxButton: document.getElementById("testMailboxButton"),
    deleteMailboxButton: document.getElementById("deleteMailboxButton"),
    importDrawer: document.getElementById("importDrawer"),
    importForm: document.getElementById("importForm"),
    importSource: document.getElementById("importSource"),
    chooseImportFileButton: document.getElementById("chooseImportFileButton"),
    importFileInput: document.getElementById("importFileInput"),
    importFileLabel: document.getElementById("importFileLabel"),
    closeImportButton: document.getElementById("closeImportButton"),
    resetImportButton: document.getElementById("resetImportButton"),
    submitImportButton: document.getElementById("submitImportButton"),
    searchInput: document.getElementById("searchInput"),
    refreshMessagesButton: document.getElementById("refreshMessagesButton"),
    overviewStrip: document.getElementById("overviewStrip"),
    overviewCards: document.getElementById("overviewCards"),
    overviewCardTemplate: document.getElementById("overviewCardTemplate"),
    folderList: document.getElementById("folderList"),
    folderSummary: document.getElementById("folderSummary"),
    readStateSelect: document.getElementById("readStateSelect"),
    importanceSelect: document.getElementById("importanceSelect"),
    sortOrderSelect: document.getElementById("sortOrderSelect"),
    pageSizeSelect: document.getElementById("pageSizeSelect"),
    attachmentOnlyCheckbox: document.getElementById("attachmentOnlyCheckbox"),
    flaggedOnlyCheckbox: document.getElementById("flaggedOnlyCheckbox"),
    selectedMessageCount: document.getElementById("selectedMessageCount"),
    clearSelectedMessagesButton: document.getElementById("clearSelectedMessagesButton"),
    batchMarkReadButton: document.getElementById("batchMarkReadButton"),
    batchMarkUnreadButton: document.getElementById("batchMarkUnreadButton"),
    batchFlagButton: document.getElementById("batchFlagButton"),
    batchUnflagButton: document.getElementById("batchUnflagButton"),
    batchMoveFolderSelect: document.getElementById("batchMoveFolderSelect"),
    batchMoveButton: document.getElementById("batchMoveButton"),
    batchArchiveButton: document.getElementById("batchArchiveButton"),
    batchDeleteButton: document.getElementById("batchDeleteButton"),
    messageBatchPanel: document.getElementById("messageBatchPanel"),
    messageList: document.getElementById("messageList"),
    messageItemTemplate: document.getElementById("messageItemTemplate"),
    messageCount: document.getElementById("messageCount"),
    messagePrevPageButton: document.getElementById("messagePrevPageButton"),
    messageNextPageButton: document.getElementById("messageNextPageButton"),
    messagePageInfo: document.getElementById("messagePageInfo"),
    listHint: document.getElementById("listHint"),
    detailEmpty: document.getElementById("detailEmpty"),
    detailCard: document.getElementById("detailCard"),
    detailMethodTag: document.getElementById("detailMethodTag"),
    detailSubject: document.getElementById("detailSubject"),
    detailSender: document.getElementById("detailSender"),
    detailTime: document.getElementById("detailTime"),
    detailTo: document.getElementById("detailTo"),
    detailCc: document.getElementById("detailCc"),
    detailBcc: document.getElementById("detailBcc"),
    detailMessageId: document.getElementById("detailMessageId"),
    detailConversationId: document.getElementById("detailConversationId"),
    detailReadBadge: document.getElementById("detailReadBadge"),
    detailFlagBadge: document.getElementById("detailFlagBadge"),
    detailImportanceBadge: document.getElementById("detailImportanceBadge"),
    detailAttachmentBadge: document.getElementById("detailAttachmentBadge"),
    detailAttachmentsSection: document.getElementById("detailAttachmentsSection"),
    detailAttachmentList: document.getElementById("detailAttachmentList"),
    detailBody: document.getElementById("detailBody"),
    detailHtmlFrame: document.getElementById("detailHtmlFrame"),
    detailHeaders: document.getElementById("detailHeaders"),
    detailTextViewButton: document.getElementById("detailTextViewButton"),
    detailHtmlViewButton: document.getElementById("detailHtmlViewButton"),
    detailHeadersViewButton: document.getElementById("detailHeadersViewButton"),
    markReadButton: document.getElementById("markReadButton"),
    markUnreadButton: document.getElementById("markUnreadButton"),
    flagMessageButton: document.getElementById("flagMessageButton"),
    unflagMessageButton: document.getElementById("unflagMessageButton"),
    moveFolderSelect: document.getElementById("moveFolderSelect"),
    moveMessageButton: document.getElementById("moveMessageButton"),
    archiveMessageButton: document.getElementById("archiveMessageButton"),
    deleteMessageButton: document.getElementById("deleteMessageButton"),
    prevMessageButton: document.getElementById("prevMessageButton"),
    nextMessageButton: document.getElementById("nextMessageButton"),
};

document.addEventListener("DOMContentLoaded", () => {
    hydrateWorkspaceState();
    bindEvents();
    checkAuth();
});

function bindEvents() {
    elements.loginForm.addEventListener("submit", async (event) => {
        event.preventDefault();
        await login();
    });

    elements.logoutButton.addEventListener("click", async () => {
        await logout();
    });

    elements.newMailboxButton.addEventListener("click", () => {
        openEditor("create");
        setStatus("开始创建新的邮箱档案");
    });

    elements.mailboxSearchInput.addEventListener("input", () => {
        state.mailboxSearchTerm = readString(elements.mailboxSearchInput.value);
        state.mailboxPage = 1;
        persistWorkspaceState();
        window.clearTimeout(state.mailboxSearchTimer);
        state.mailboxSearchTimer = window.setTimeout(() => {
            loadMailboxes();
        }, 250);
    });

    elements.mailboxPrevPageButton.addEventListener("click", async () => {
        if (state.mailboxPage <= 1) {
            return;
        }
        state.mailboxPage -= 1;
        persistWorkspaceState();
        await loadMailboxes();
    });

    elements.mailboxNextPageButton.addEventListener("click", async () => {
        if (!state.mailboxMeta.has_next) {
            return;
        }
        state.mailboxPage += 1;
        persistWorkspaceState();
        await loadMailboxes();
    });

    elements.toggleVisibleMailboxesButton.addEventListener("click", () => {
        toggleVisibleMailboxesSelection();
    });

    elements.clearSelectedMailboxesButton.addEventListener("click", () => {
        clearSelectedMailboxes();
    });

    elements.testVisibleMailboxesButton.addEventListener("click", async () => {
        await testVisibleMailboxes();
    });

    elements.testSelectedMailboxesButton.addEventListener("click", async () => {
        await testSelectedMailboxes();
    });

    elements.bulkPreferredMethodSelect.addEventListener("change", () => {
        state.bulkPreferredMethod = readString(elements.bulkPreferredMethodSelect.value);
        persistWorkspaceState();
        renderBulkActions();
    });

    elements.applyPreferredMethodButton.addEventListener("click", async () => {
        await applyBulkPreferredMethod();
    });

    elements.deleteSelectedMailboxesButton.addEventListener("click", async () => {
        await deleteSelectedMailboxes();
    });

    elements.importMailboxesButton.addEventListener("click", () => {
        openImportDrawer();
        setStatus("请粘贴批量导入文本");
    });

    elements.closeEditorButton.addEventListener("click", () => {
        closeEditor();
    });

    elements.closeImportButton.addEventListener("click", () => {
        closeImportDrawer();
    });

    elements.chooseImportFileButton.addEventListener("click", () => {
        elements.importFileInput.click();
    });

    elements.importFileInput.addEventListener("change", async (event) => {
        await loadImportFile(event);
    });

    elements.mailboxForm.addEventListener("submit", async (event) => {
        event.preventDefault();
        await saveMailbox();
    });

    elements.importForm.addEventListener("submit", async (event) => {
        event.preventDefault();
        await importMailboxes();
    });

    elements.resetMailboxButton.addEventListener("click", () => {
        if (state.editorMode === "edit" && state.selectedMailboxDetail) {
            fillMailboxForm(state.selectedMailboxDetail);
            setStatus("已恢复当前邮箱配置");
            return;
        }
        resetMailboxForm();
        setStatus("表单已重置");
    });

    elements.testMailboxButton.addEventListener("click", async () => {
        await testMailboxConnection();
    });

    elements.resetImportButton.addEventListener("click", () => {
        resetImportForm();
        setStatus("批量导入文本已清空");
    });

    elements.deleteMailboxButton.addEventListener("click", async () => {
        if (!state.selectedMailboxId) {
            return;
        }
        if (!window.confirm("确认删除这个邮箱档案？")) {
            return;
        }
        await deleteMailbox(state.selectedMailboxId);
    });

    elements.searchInput.addEventListener("input", () => {
        state.searchTerm = elements.searchInput.value.trim();
        state.messagePage = 1;
        window.clearTimeout(state.searchTimer);
        state.searchTimer = window.setTimeout(() => {
            if (state.selectedMailboxId) {
                loadMessages({ silent: true, keepSelection: false });
            }
        }, 320);
    });

    elements.refreshMessagesButton.addEventListener("click", async () => {
        await loadFolders({ silent: true });
        await loadMessages({ silent: true, keepSelection: true });
    });

    elements.readStateSelect.addEventListener("change", async () => {
        state.readState = readString(elements.readStateSelect.value) || "all";
        state.messagePage = 1;
        await loadMessages({ silent: true, keepSelection: false });
    });

    elements.importanceSelect.addEventListener("change", async () => {
        state.importance = readString(elements.importanceSelect.value) || "all";
        state.messagePage = 1;
        await loadMessages({ silent: true, keepSelection: false });
    });

    elements.sortOrderSelect.addEventListener("change", async () => {
        state.sortOrder = readString(elements.sortOrderSelect.value) || "desc";
        state.messagePage = 1;
        await loadMessages({ silent: true, keepSelection: false });
    });

    elements.pageSizeSelect.addEventListener("change", async () => {
        state.messagePageSize = Number.parseInt(elements.pageSizeSelect.value, 10) || 20;
        state.messagePage = 1;
        await loadMessages({ silent: true, keepSelection: false });
    });

    elements.attachmentOnlyCheckbox.addEventListener("change", async () => {
        state.hasAttachmentsOnly = elements.attachmentOnlyCheckbox.checked;
        state.messagePage = 1;
        await loadMessages({ silent: true, keepSelection: false });
    });

    elements.flaggedOnlyCheckbox.addEventListener("change", async () => {
        state.flaggedOnly = elements.flaggedOnlyCheckbox.checked;
        state.messagePage = 1;
        await loadMessages({ silent: true, keepSelection: false });
    });

    elements.messagePrevPageButton.addEventListener("click", async () => {
        if (state.messagePage <= 1) {
            return;
        }
        state.messagePage -= 1;
        await loadMessages({ silent: true, keepSelection: false });
    });

    elements.messageNextPageButton.addEventListener("click", async () => {
        if (!state.messageMeta.has_next) {
            return;
        }
        state.messagePage += 1;
        await loadMessages({ silent: true, keepSelection: false });
    });

    elements.clearSelectedMessagesButton.addEventListener("click", () => {
        state.selectedMessageIds = new Set();
        renderMessageBatchActions();
        renderMessages();
    });

    elements.batchMarkReadButton.addEventListener("click", async () => {
        await runBatchMessageAction("mark_read");
    });

    elements.batchMarkUnreadButton.addEventListener("click", async () => {
        await runBatchMessageAction("mark_unread");
    });

    elements.batchFlagButton.addEventListener("click", async () => {
        await runBatchMessageAction("flag");
    });

    elements.batchUnflagButton.addEventListener("click", async () => {
        await runBatchMessageAction("unflag");
    });

    elements.batchMoveButton.addEventListener("click", async () => {
        await runBatchMessageAction("move", readString(elements.batchMoveFolderSelect.value));
    });

    elements.batchArchiveButton.addEventListener("click", async () => {
        await runBatchMessageAction("archive");
    });

    elements.batchDeleteButton.addEventListener("click", async () => {
        await runBatchMessageAction("delete");
    });

    elements.markReadButton.addEventListener("click", async () => {
        await updateReadState(true);
    });

    elements.markUnreadButton.addEventListener("click", async () => {
        await updateReadState(false);
    });

    elements.flagMessageButton.addEventListener("click", async () => {
        await updateFlagState(true);
    });

    elements.unflagMessageButton.addEventListener("click", async () => {
        await updateFlagState(false);
    });

    elements.moveMessageButton.addEventListener("click", async () => {
        const targetFolder = readString(elements.moveFolderSelect.value);
        if (!targetFolder) {
            setStatus("请先选择目标文件夹", true);
            return;
        }
        await moveSelectedMessage(targetFolder);
    });

    elements.archiveMessageButton.addEventListener("click", async () => {
        await moveSelectedMessage("archive");
    });

    elements.deleteMessageButton.addEventListener("click", async () => {
        await deleteSelectedMessage();
    });

    elements.detailTextViewButton.addEventListener("click", () => {
        state.detailView = "text";
        renderDetail(state.selectedMessage);
    });

    elements.detailHtmlViewButton.addEventListener("click", () => {
        state.detailView = "html";
        renderDetail(state.selectedMessage);
    });

    elements.detailHeadersViewButton.addEventListener("click", () => {
        state.detailView = "headers";
        renderDetail(state.selectedMessage);
    });

    elements.prevMessageButton.addEventListener("click", async () => {
        await selectAdjacentMessage(-1);
    });

    elements.nextMessageButton.addEventListener("click", async () => {
        await selectAdjacentMessage(1);
    });

    document.addEventListener("visibilitychange", () => {
        if (document.hidden) {
            stopAutoRefresh();
            return;
        }
        startAutoRefresh({ immediate: true });
    });
}

async function checkAuth() {
    try {
        const payload = await requestJson("/api/auth/me");
        applyAuthState(payload);
        if (payload.authenticated) {
            await loadMailboxes();
        } else {
            resetWorkspace();
        }
    } catch (error) {
        applyAuthState({ authenticated: false });
        setStatus(error.message, true);
    }
}

async function login() {
    const formData = new FormData(elements.loginForm);
    setBusy([elements.loginButton], true);

    try {
        const payload = await requestJson("/api/auth/login", {
            method: "POST",
            body: {
                username: readString(formData.get("username")),
                password: readString(formData.get("password")),
            },
        });
        applyAuthState(payload);
        await loadMailboxes();
        setStatus("管理员登录成功");
    } catch (error) {
        setStatus(error.message, true);
    } finally {
        setBusy([elements.loginButton], false);
    }
}

async function logout() {
    try {
        await requestJson("/api/auth/logout", { method: "POST" });
        applyAuthState({ authenticated: false });
        resetWorkspace();
        setStatus("已退出登录");
    } catch (error) {
        setStatus(error.message, true);
    }
}

async function loadMailboxes() {
    const currentToken = ++state.mailboxFetchToken;
    const params = new URLSearchParams({
        page: String(state.mailboxPage),
        page_size: String(state.mailboxPageSize),
    });
    if (state.mailboxSearchTerm) {
        params.set("q", state.mailboxSearchTerm);
    }

    try {
        const payload = await requestJson(`/api/mailboxes?${params.toString()}`);
        if (currentToken !== state.mailboxFetchToken) {
            return;
        }

        state.mailboxes = payload.items || [];
        state.mailboxMeta = payload.meta || createDefaultMailboxMeta();
        state.mailboxPage = state.mailboxMeta.page || state.mailboxPage;
        persistWorkspaceState();
        if (!state.mailboxes.length && state.mailboxMeta.total_pages > 0 && state.mailboxPage > state.mailboxMeta.total_pages) {
            state.mailboxPage = state.mailboxMeta.total_pages;
            persistWorkspaceState();
            await loadMailboxes();
            return;
        }
        renderMailboxList();

        if (!state.mailboxSearchTerm && state.mailboxMeta.total === 0) {
            resetWorkspace({ preserveAuth: true });
            openEditor("create");
            setStatus("当前还没有邮箱档案，请先创建一个");
            return;
        }

        const currentMailbox = state.mailboxes.find((item) => item.id === state.selectedMailboxId) || null;
        if (currentMailbox) {
            state.selectedMailbox = currentMailbox;
            elements.currentMailboxLabel.textContent = currentMailbox.label;
            renderMailboxList();
            return;
        }

        if (!state.selectedMailboxId && state.mailboxes.length) {
            await selectMailbox(state.mailboxes[0].id);
        }
    } catch (error) {
        if (currentToken !== state.mailboxFetchToken) {
            return;
        }
        setStatus(error.message, true);
    }
}

async function selectMailbox(mailboxId) {
    const mailbox = state.mailboxes.find((item) => item.id === mailboxId) || null;
    state.messageFetchToken += 1;
    state.detailFetchToken += 1;
    state.selectedMailboxId = mailbox?.id || null;
    state.selectedMailbox = mailbox;
    state.selectedMailboxDetail = null;
    state.activeMethod = mailbox?.preferred_method || "graph_api";
    state.folders = [];
    state.selectedFolder = "INBOX";
    state.overviewMethods = [];
    state.messageMeta = createDefaultMessageMeta();
    state.messagePage = 1;
    state.selectedMessageIds = new Set();
    state.selectedMessage = null;
    state.selectedMessageId = null;
    state.messages = [];
    elements.currentMailboxLabel.textContent = mailbox ? mailbox.label : "未选择";
    elements.currentMethodLabel.textContent = mailbox ? labelForMethod(state.activeMethod) : "未设置";
    renderMailboxList();
    renderOverview([]);
    renderFolderList();
    renderMessageBatchActions();
    renderMessages();
    renderDetail(null);

    if (!mailbox) {
        stopAutoRefresh();
        return;
    }

    closeEditor();
    setStatus(`已切换到邮箱：${mailbox.label}`);
    await loadOverview();
    await loadFolders({ silent: true });
    await loadMessages({ silent: true, keepSelection: false });
    startAutoRefresh();
}

function openEditor(mode, mailbox = null) {
    closeImportDrawer();
    state.editorMode = mode;
    elements.editorModeLabel.textContent = mode === "edit" ? "编辑邮箱" : "新增邮箱";
    elements.editorDrawer.hidden = false;
    elements.deleteMailboxButton.disabled = mode !== "edit" || !mailbox;

    if (mode === "edit" && mailbox) {
        state.selectedMailboxDetail = mailbox;
        fillMailboxForm(mailbox);
        return;
    }
    state.selectedMailboxDetail = null;
    resetMailboxForm();
}

function closeEditor() {
    elements.editorDrawer.hidden = true;
}

function openImportDrawer() {
    closeEditor();
    elements.importDrawer.hidden = false;
    if (!readString(elements.importSource.value)) {
        elements.importForm.elements.preferred_method.value = "graph_api";
    }
}

function closeImportDrawer() {
    elements.importDrawer.hidden = true;
}

async function loadMailboxDetail(mailboxId) {
    const payload = await requestJson(`/api/mailboxes/${mailboxId}`);
    return payload.mailbox;
}

async function saveMailbox() {
    const payload = extractMailboxPayload();
    const mailboxId = readString(elements.mailboxForm.elements.mailbox_id.value);
    const isUpdate = Boolean(mailboxId);

    setBusy([elements.saveMailboxButton, elements.deleteMailboxButton], true);
    try {
        const response = await requestJson(
            isUpdate ? `/api/mailboxes/${mailboxId}` : "/api/mailboxes",
            {
                method: isUpdate ? "PUT" : "POST",
                body: payload,
            }
        );
        revealMailboxInSidebar(response.mailbox.id);
        await loadMailboxes();
        closeEditor();
        await selectMailbox(response.mailbox.id);
        setStatus(isUpdate ? "邮箱档案已更新" : "邮箱档案已创建");
    } catch (error) {
        setStatus(error.message, true);
    } finally {
        setBusy([elements.saveMailboxButton, elements.deleteMailboxButton], false);
    }
}

async function deleteMailbox(mailboxId) {
    setBusy([elements.deleteMailboxButton], true);
    try {
        await requestJson(`/api/mailboxes/${mailboxId}`, { method: "DELETE" });
        const nextSelected = new Set(state.selectedMailboxIds);
        nextSelected.delete(mailboxId);
        state.selectedMailboxIds = nextSelected;
        persistWorkspaceState();
        stopAutoRefresh();
        state.selectedMailboxId = null;
        state.selectedMailbox = null;
        state.selectedMailboxDetail = null;
        closeEditor();
        await loadMailboxes();
        setStatus("邮箱档案已删除");
    } catch (error) {
        setStatus(error.message, true);
    } finally {
        setBusy([elements.deleteMailboxButton], false);
    }
}

async function importMailboxes() {
    const payload = {
        raw_text: readString(elements.importForm.elements.raw_text.value),
        preferred_method: readString(elements.importForm.elements.preferred_method.value) || "graph_api",
    };

    setBusy([elements.submitImportButton, elements.resetImportButton, elements.closeImportButton], true);
    try {
        const response = await requestJson("/api/mailboxes/import", {
            method: "POST",
            body: payload,
        });

        const importedMailboxes = Array.isArray(response.mailboxes) ? response.mailboxes : [];
        const importedIds = importedMailboxes
            .map((item) => item?.id)
            .filter((value) => Number.isInteger(value));
        const targetMailboxId = importedIds[importedIds.length - 1] || state.selectedMailboxId;
        if (targetMailboxId) {
            revealMailboxInSidebar(targetMailboxId);
        }
        await loadMailboxes();
        if (targetMailboxId) {
            await selectMailbox(targetMailboxId);
        }

        closeImportDrawer();
        resetImportForm();
        setStatus(formatImportSummary(response.summary));
    } catch (error) {
        setStatus(error.message, true);
    } finally {
        setBusy([elements.submitImportButton, elements.resetImportButton, elements.closeImportButton], false);
    }
}

async function loadImportFile(event) {
    const input = event.target;
    const [file] = input?.files || [];
    if (!file) {
        setImportFileLabel("");
        return;
    }

    try {
        const content = await file.text();
        const trimmed = readString(content);
        if (!trimmed) {
            throw new Error("所选文件为空");
        }

        elements.importSource.value = content;
        setImportFileLabel(file.name);
        setStatus(`已载入导入文件：${file.name}`);
    } catch (error) {
        input.value = "";
        setImportFileLabel("");
        setStatus(error.message || "读取导入文件失败", true);
    }
}

async function loadOverview() {
    if (!ensureMailboxSelected()) {
        return;
    }
    try {
        const payload = await requestJson("/api/mailbox/overview", {
            method: "POST",
            body: { mailbox_id: state.selectedMailboxId },
        });
        state.overviewMethods = payload.overview.methods || [];
        renderOverview(state.overviewMethods);
    } catch (error) {
        state.overviewMethods = [];
        renderOverview([]);
        setStatus(error.message, true);
    }
}

async function loadFolders({ silent = false } = {}) {
    if (!ensureMailboxSelected()) {
        return;
    }
    try {
        const payload = await requestJson("/api/mailbox/folders", {
            method: "POST",
            body: {
                mailbox_id: state.selectedMailboxId,
                method: state.activeMethod || state.selectedMailbox?.preferred_method || "graph_api",
            },
        });
        state.folders = Array.isArray(payload.folders) ? payload.folders : [];
        const nextFolder = state.folders.find((item) => item.id === state.selectedFolder)
            ? state.selectedFolder
            : (state.folders.find((item) => item.is_default)?.id || state.folders[0]?.id || "INBOX");
        state.selectedFolder = nextFolder;
        renderFolderList();
        syncFolderSelectOptions();
    } catch (error) {
        state.folders = [];
        renderFolderList();
        syncFolderSelectOptions();
        if (!silent) {
            setStatus(error.message, true);
        }
    }
}

async function loadMessages({ silent = false, keepSelection = true } = {}) {
    if (!ensureMailboxSelected() || state.isLoadingMessages) {
        return;
    }

    state.isLoadingMessages = true;
    const currentToken = ++state.messageFetchToken;
    const method = state.activeMethod || state.selectedMailbox?.preferred_method || "graph_api";
    if (!silent) {
        setBusy([elements.messageList], true);
    }

    try {
        const payload = await requestJson("/api/mailbox/messages", {
            method: "POST",
            body: {
                mailbox_id: state.selectedMailboxId,
                method,
                folder: state.selectedFolder || "INBOX",
                top: state.messagePageSize,
                page: state.messagePage,
                page_size: state.messagePageSize,
                unread_only: state.readState === "unread",
                read_state: state.readState,
                has_attachments_only: state.hasAttachmentsOnly,
                flagged_only: state.flaggedOnly,
                importance: state.importance,
                sort_order: state.sortOrder,
                keyword: state.searchTerm,
            },
        });

        if (currentToken !== state.messageFetchToken) {
            return;
        }

        state.activeMethod = payload.method || method;
        state.selectedFolder = payload.folder || state.selectedFolder;
        state.messageMeta = payload.meta || createDefaultMessageMeta();
        const previousSignature = buildMessageListSignature(state.messages);
        const previousSelectedMessageId = state.selectedMessageId;
        const nextMessages = payload.messages || [];
        const nextSelectedMessageId = keepSelection && previousSelectedMessageId && nextMessages.some((item) => item.message_id === previousSelectedMessageId)
            ? previousSelectedMessageId
            : nextMessages[0]?.message_id || null;

        state.messages = nextMessages;
        state.selectedMessageId = nextSelectedMessageId;

        if (buildMessageListSignature(nextMessages) !== previousSignature || nextSelectedMessageId !== previousSelectedMessageId) {
            renderMessages();
        } else {
            updateMessageSelection();
        }

        setElementText(elements.messageCount, `${payload.count} 封邮件`);
        setElementText(
            elements.listHint,
            state.searchTerm
                ? `搜索：${state.searchTerm}`
                : `${state.selectedFolder || "INBOX"} · ${labelForMethod(state.activeMethod)}`
        );
        setElementText(elements.currentMethodLabel, labelForMethod(state.activeMethod));
        setElementText(elements.lastSyncLabel, formatDate(new Date().toISOString(), true));
        renderFolderList();
        renderMessageBatchActions();
        renderOverview(state.overviewMethods);
        renderMessagePagination();

        if (state.selectedMessageId) {
            const selectedSummary = findMessageById(state.selectedMessageId);
            if (selectedSummary && shouldRefetchDetail(selectedSummary)) {
                await loadMessageDetail(state.selectedMessageId, { silent: true });
            } else if (selectedSummary) {
                state.selectedMessage = mergeMessageSummary(state.selectedMessage, selectedSummary);
                renderDetail(state.selectedMessage);
            }
        } else {
            state.selectedMessage = null;
            state.detailFetchToken += 1;
            renderDetail(null);
            if (!silent) {
                setStatus("当前筛选条件下没有邮件");
            }
        }

        if (!silent) {
            setStatus("邮件列表已更新");
        }
    } catch (error) {
        state.messages = [];
        renderOverview(state.overviewMethods);
        renderMessages();
        renderDetail(null);
        setElementText(elements.messageCount, "0 封邮件");
        state.messageMeta = createDefaultMessageMeta();
        renderMessagePagination();
        setElementText(elements.listHint, "加载失败");
        setStatus(error.message, true);
    } finally {
        state.isLoadingMessages = false;
        if (!silent) {
            setBusy([elements.messageList], false);
        }
    }
}

async function loadMessageDetail(messageId, { silent = false } = {}) {
    if (!ensureMailboxSelected()) {
        return;
    }

    state.selectedMessageId = messageId;
    state.selectedMessage = mergeMessageSummary(
        state.selectedMessage?.message_id === messageId ? state.selectedMessage : null,
        findMessageById(messageId)
    );
    const currentToken = ++state.detailFetchToken;
    updateMessageSelection();
    renderDetail(state.selectedMessage);

    try {
        const payload = await requestJson("/api/mailbox/message", {
            method: "POST",
            body: {
                mailbox_id: state.selectedMailboxId,
                method: state.activeMethod || state.selectedMailbox?.preferred_method || "graph_api",
                folder: state.selectedFolder || "INBOX",
                message_id: messageId,
            },
        });
        if (currentToken !== state.detailFetchToken || messageId !== state.selectedMessageId) {
            return;
        }
        state.selectedMessage = mergeMessageSummary(payload.message, findMessageById(messageId));
        renderDetail(state.selectedMessage);
        if (!silent) {
            setStatus(`已打开邮件：${payload.message.subject || "无主题"}`);
        }
    } catch (error) {
        state.selectedMessage = mergeMessageSummary(null, findMessageById(messageId));
        renderDetail(state.selectedMessage);
        setStatus(error.message, true);
    }
}

async function updateReadState(isRead) {
    if (!ensureMailboxSelected() || !state.selectedMessage) {
        setStatus("请先选择一封邮件", true);
        return;
    }

    setBusy([elements.markReadButton, elements.markUnreadButton], true);
    try {
        const payload = await requestJson("/api/mailbox/message/read-state", {
            method: "POST",
            body: {
                mailbox_id: state.selectedMailboxId,
                method: state.activeMethod || state.selectedMailbox?.preferred_method || "graph_api",
                folder: state.selectedFolder || "INBOX",
                message_id: state.selectedMessage.message_id,
                is_read: isRead,
            },
        });
        state.selectedMessage = payload.message;
        state.messages = state.messages.map((item) =>
            item.message_id === payload.message.message_id
                ? { ...item, is_read: payload.message.is_read }
                : item
        );
        renderMessages();
        renderDetail(payload.message);
        setStatus(isRead ? "邮件已标记为已读" : "邮件已标记为未读");
    } catch (error) {
        setStatus(error.message, true);
    } finally {
        setBusy([elements.markReadButton, elements.markUnreadButton], false);
    }
}

async function updateFlagState(isFlagged) {
    if (!ensureMailboxSelected() || !state.selectedMessage) {
        setStatus("请先选择一封邮件", true);
        return;
    }

    setBusy([elements.flagMessageButton, elements.unflagMessageButton], true);
    try {
        const payload = await requestJson("/api/mailbox/message/flag-state", {
            method: "POST",
            body: {
                mailbox_id: state.selectedMailboxId,
                method: state.activeMethod || state.selectedMailbox?.preferred_method || "graph_api",
                folder: state.selectedFolder || "INBOX",
                message_id: state.selectedMessage.message_id,
                is_flagged: isFlagged,
            },
        });
        state.selectedMessage = payload.message;
        state.messages = state.messages.map((item) =>
            item.message_id === payload.message.message_id
                ? { ...item, is_flagged: payload.message.is_flagged }
                : item
        );
        renderMessages();
        renderDetail(payload.message);
        setStatus(isFlagged ? "邮件已加星" : "邮件已取消星标");
    } catch (error) {
        setStatus(error.message, true);
    } finally {
        setBusy([elements.flagMessageButton, elements.unflagMessageButton], false);
    }
}

async function moveSelectedMessage(destinationFolder) {
    if (!ensureMailboxSelected() || !state.selectedMessage) {
        setStatus("请先选择一封邮件", true);
        return;
    }

    setBusy([elements.moveMessageButton, elements.archiveMessageButton], true);
    try {
        await requestJson("/api/mailbox/message/move", {
            method: "POST",
            body: {
                mailbox_id: state.selectedMailboxId,
                method: state.activeMethod || state.selectedMailbox?.preferred_method || "graph_api",
                folder: state.selectedFolder || "INBOX",
                message_id: state.selectedMessage.message_id,
                destination_folder: destinationFolder,
            },
        });
        state.selectedMessage = null;
        state.selectedMessageId = null;
        await loadFolders({ silent: true });
        await loadMessages({ silent: true, keepSelection: false });
        renderDetail(null);
        setStatus(`邮件已移动到 ${destinationFolder}`);
    } catch (error) {
        setStatus(error.message, true);
    } finally {
        setBusy([elements.moveMessageButton, elements.archiveMessageButton], false);
    }
}

async function deleteSelectedMessage() {
    if (!ensureMailboxSelected() || !state.selectedMessage) {
        setStatus("请先选择一封邮件", true);
        return;
    }

    setBusy([elements.deleteMessageButton], true);
    try {
        await requestJson("/api/mailbox/message/delete", {
            method: "POST",
            body: {
                mailbox_id: state.selectedMailboxId,
                method: state.activeMethod || state.selectedMailbox?.preferred_method || "graph_api",
                folder: state.selectedFolder || "INBOX",
                message_id: state.selectedMessage.message_id,
            },
        });
        state.selectedMessage = null;
        state.selectedMessageId = null;
        await loadFolders({ silent: true });
        await loadMessages({ silent: true, keepSelection: false });
        renderDetail(null);
        setStatus("邮件已删除");
    } catch (error) {
        setStatus(error.message, true);
    } finally {
        setBusy([elements.deleteMessageButton], false);
    }
}

function startAutoRefresh({ immediate = false } = {}) {
    stopAutoRefresh();
    if (!state.authenticated || !state.selectedMailboxId || document.hidden) {
        return;
    }
    state.autoRefreshTimer = window.setTimeout(async () => {
        state.autoRefreshTimer = null;
        await loadMessages({ silent: true, keepSelection: true });
        startAutoRefresh();
    }, immediate ? 0 : AUTO_REFRESH_MS);
}

function stopAutoRefresh() {
    if (state.autoRefreshTimer) {
        window.clearTimeout(state.autoRefreshTimer);
        state.autoRefreshTimer = null;
    }
}

function applyAuthState(payload) {
    state.authenticated = Boolean(payload.authenticated);
    state.username = payload.username || "";
    elements.authView.hidden = state.authenticated;
    elements.dashboardView.hidden = !state.authenticated;
    elements.logoutButton.hidden = !state.authenticated;
    elements.sessionLabel.textContent = state.authenticated ? state.username : "未登录";
    if (!state.authenticated) {
        elements.currentMailboxLabel.textContent = "未选择";
        elements.currentMethodLabel.textContent = "Graph API";
        stopAutoRefresh();
    }
}

function renderMailboxList() {
    elements.mailboxList.innerHTML = "";
    if (!state.mailboxes.length) {
        elements.mailboxList.innerHTML = state.mailboxSearchTerm
            ? "<div class='detail-empty compact-empty'>没有匹配的邮箱档案</div>"
            : "<div class='detail-empty compact-empty'>还没有邮箱档案</div>";
        renderSidebarFillState();
        renderMailboxPagination({
            selectedMailboxVisible: false,
        });
        renderBulkActions();
        return;
    }

    const selectedMailboxVisible = state.mailboxes.some((item) => item.id === state.selectedMailboxId);
    state.mailboxes.forEach((mailbox) => {
        const fragment = elements.mailboxItemTemplate.content.cloneNode(true);
        const article = fragment.querySelector(".mailbox-card");
        const checkbox = fragment.querySelector(".mailbox-select-checkbox");
        const mainButton = fragment.querySelector(".mailbox-main");
        const editButton = fragment.querySelector(".mailbox-edit");
        const statusText = fragment.querySelector(".mailbox-item-status");
        const connectionResult = state.mailboxConnectionResults[mailbox.id] || null;
        const isChecked = state.selectedMailboxIds.has(mailbox.id);

        article.classList.toggle("is-active", mailbox.id === state.selectedMailboxId);
        article.classList.toggle("is-checked", isChecked);
        mainButton.querySelector(".mailbox-item-label").textContent = mailbox.label;
        mainButton.querySelector(".mailbox-item-email").textContent = mailbox.email;
        mainButton.querySelector(".mailbox-item-method").textContent = labelForMethod(mailbox.preferred_method);
        if (checkbox) {
            checkbox.checked = isChecked;
            checkbox.addEventListener("click", (event) => {
                event.stopPropagation();
            });
            checkbox.addEventListener("change", (event) => {
                event.stopPropagation();
                setMailboxSelection(mailbox.id, checkbox.checked);
            });
        }
        if (statusText) {
            if (connectionResult?.message) {
                statusText.hidden = false;
                statusText.textContent = connectionResult.message;
                statusText.classList.toggle("is-success", Boolean(connectionResult.success));
                statusText.classList.toggle("is-error", !connectionResult.success);
            } else {
                statusText.hidden = true;
                statusText.textContent = "";
                statusText.classList.remove("is-success", "is-error");
            }
        }

        mainButton.addEventListener("click", async () => {
            await selectMailbox(mailbox.id);
        });

        editButton.addEventListener("click", async (event) => {
            event.stopPropagation();
            setBusy([editButton], true);
            try {
                await selectMailbox(mailbox.id);
                const detail = await loadMailboxDetail(mailbox.id);
                openEditor("edit", detail);
                setStatus(`正在编辑邮箱：${detail.label}`);
            } catch (error) {
                setStatus(error.message, true);
            } finally {
                setBusy([editButton], false);
            }
        });

        elements.mailboxList.append(fragment);
    });

    renderMailboxPagination({
        selectedMailboxVisible,
    });
    renderSidebarFillState();
    renderBulkActions();
}

function renderSidebarFillState() {
    const shouldShow = state.mailboxes.length <= 2;
    elements.sidebarFillState.hidden = !shouldShow;
}

async function testVisibleMailboxes() {
    if (!state.mailboxes.length) {
        setStatus("当前页没有可测试的邮箱", true);
        return;
    }

    await runMailboxBatchConnectionTest({
        mailboxIds: state.mailboxes.map((item) => item.id),
        busyTargets: [elements.testVisibleMailboxesButton],
        successLabel: "本页测试完成",
    });
}

async function testSelectedMailboxes() {
    const mailboxIds = getSelectedMailboxIds();
    if (!mailboxIds.length) {
        setStatus("请先选择要测试的邮箱", true);
        return;
    }

    await runMailboxBatchConnectionTest({
        mailboxIds,
        busyTargets: [elements.testSelectedMailboxesButton],
        successLabel: "已选邮箱测试完成",
    });
}

async function runMailboxBatchConnectionTest({ mailboxIds, busyTargets, successLabel }) {
    setBusy(busyTargets, true);
    try {
        const payload = await requestJson("/api/mailboxes/test-connection/batch", {
            method: "POST",
            body: {
                mailbox_ids: mailboxIds,
            },
        });

        const nextResults = { ...state.mailboxConnectionResults };
        (payload.results || []).forEach((item) => {
            if (!Number.isInteger(item?.mailbox_id)) {
                return;
            }
            nextResults[item.mailbox_id] = {
                success: Boolean(item.success),
                method: item.method || "",
                message: item.message || (item.success ? "连接成功" : "连接失败"),
            };
        });
        state.mailboxConnectionResults = nextResults;
        renderMailboxList();

        const summary = payload.summary || {};
        setStatus(
            `${successLabel}：成功 ${summary.succeeded || 0} 个，失败 ${summary.failed || 0} 个`
        );
    } catch (error) {
        setStatus(error.message, true);
    } finally {
        setBusy(busyTargets, false);
    }
}

async function applyBulkPreferredMethod() {
    const mailboxIds = getSelectedMailboxIds();
    if (!mailboxIds.length) {
        setStatus("请先选择要更新的邮箱", true);
        return;
    }

    const targetMethod = state.bulkPreferredMethod;
    if (!targetMethod) {
        setStatus("请先选择要切换到的默认方法", true);
        return;
    }

    setBusy([elements.applyPreferredMethodButton, elements.bulkPreferredMethodSelect], true);
    try {
        const payload = await requestJson("/api/mailboxes/preferred-method/batch", {
            method: "POST",
            body: {
                mailbox_ids: mailboxIds,
                preferred_method: targetMethod,
            },
        });

        const succeededMailboxIds = (payload.results || [])
            .filter((item) => item?.success && Number.isInteger(item.mailbox_id))
            .map((item) => item.mailbox_id);
        const successfulIdSet = new Set(succeededMailboxIds);

        state.mailboxes = state.mailboxes.map((item) =>
            successfulIdSet.has(item.id)
                ? { ...item, preferred_method: targetMethod }
                : item
        );
        if (state.selectedMailbox && successfulIdSet.has(state.selectedMailbox.id)) {
            state.selectedMailbox = {
                ...state.selectedMailbox,
                preferred_method: targetMethod,
            };
        }
        if (state.selectedMailboxDetail && successfulIdSet.has(state.selectedMailboxDetail.id)) {
            state.selectedMailboxDetail = {
                ...state.selectedMailboxDetail,
                preferred_method: targetMethod,
            };
        }

        await loadMailboxes();
        renderMailboxList();

        const summary = payload.summary || {};
        setStatus(
            `批量切换完成：成功 ${summary.succeeded || 0} 个，失败 ${summary.failed || 0} 个`
        );
    } catch (error) {
        setStatus(error.message, true);
    } finally {
        setBusy([elements.applyPreferredMethodButton, elements.bulkPreferredMethodSelect], false);
    }
}

async function deleteSelectedMailboxes() {
    const mailboxIds = getSelectedMailboxIds();
    if (!mailboxIds.length) {
        setStatus("请先选择要删除的邮箱", true);
        return;
    }
    if (!window.confirm(`确认删除已选的 ${mailboxIds.length} 个邮箱档案？`)) {
        return;
    }

    setBusy([elements.deleteSelectedMailboxesButton], true);
    try {
        const payload = await requestJson("/api/mailboxes/delete/batch", {
            method: "POST",
            body: {
                mailbox_ids: mailboxIds,
            },
        });

        const succeededMailboxIds = (payload.results || [])
            .filter((item) => item?.success && Number.isInteger(item.mailbox_id))
            .map((item) => item.mailbox_id);
        const successfulIdSet = new Set(succeededMailboxIds);
        const shouldCloseEditor = Boolean(
            state.selectedMailboxDetail && successfulIdSet.has(state.selectedMailboxDetail.id)
        );

        if (successfulIdSet.size) {
            state.selectedMailboxIds = new Set(
                getSelectedMailboxIds().filter((mailboxId) => !successfulIdSet.has(mailboxId))
            );
            const nextConnectionResults = { ...state.mailboxConnectionResults };
            succeededMailboxIds.forEach((mailboxId) => {
                delete nextConnectionResults[mailboxId];
            });
            state.mailboxConnectionResults = nextConnectionResults;

            if (state.selectedMailboxId && successfulIdSet.has(state.selectedMailboxId)) {
                stopAutoRefresh();
                state.selectedMailboxId = null;
                state.selectedMailbox = null;
                state.selectedMailboxDetail = null;
                state.activeMethod = null;
                state.overviewMethods = [];
                state.messages = [];
                state.selectedMessageId = null;
                state.selectedMessage = null;
                renderOverview([]);
                renderMessages();
                renderDetail(null);
            }
        }

        persistWorkspaceState();
        await loadMailboxes();
        if (shouldCloseEditor) {
            closeEditor();
        }

        const summary = payload.summary || {};
        setStatus(
            `批量删除完成：成功 ${summary.succeeded || 0} 个，失败 ${summary.failed || 0} 个`
        );
    } catch (error) {
        setStatus(error.message, true);
    } finally {
        setBusy([elements.deleteSelectedMailboxesButton], false);
    }
}

function renderOverview(methods) {
    if (!elements.overviewStrip || !elements.overviewCards || !elements.overviewCardTemplate) {
        return;
    }

    const hasMethods = Array.isArray(methods) && methods.length > 0;
    elements.overviewCards.innerHTML = "";
    elements.overviewStrip.hidden = !hasMethods;
    if (!hasMethods) {
        return;
    }

    methods.forEach((item) => {
        const fragment = elements.overviewCardTemplate.content.cloneNode(true);
        const button = fragment.querySelector(".overview-card-mini");
        if (!button) {
            return;
        }
        button.dataset.state = item.status;
        button.classList.toggle("is-active", item.method === (state.activeMethod || state.selectedMailbox?.preferred_method || ""));
        setElementText(button.querySelector(".overview-card-method"), item.label);
        setElementText(button.querySelector(".overview-card-count"), `${item.message_count || 0} 封`);
        setElementText(button.querySelector(".overview-card-subject"), item.latest_subject || "尚未读取到邮件");
        setElementText(button.querySelector(".overview-card-message"), item.message || "连接状态待确认");
        button.addEventListener("click", async () => {
            await switchMailboxMethod(item.method);
        });
        elements.overviewCards.append(button);
    });
}

async function switchMailboxMethod(method) {
    if (!ensureMailboxSelected()) {
        return;
    }

    const targetMethod = readString(method) || state.selectedMailbox?.preferred_method || "graph_api";
    if (targetMethod === state.activeMethod && state.messages.length) {
        return;
    }

    state.activeMethod = targetMethod;
    state.selectedMessageId = null;
    state.selectedMessage = null;
    state.messages = [];
    state.selectedMessageIds = new Set();
    renderOverview(state.overviewMethods);
    renderMessageBatchActions();
    renderMessages();
    renderDetail(null);
    setStatus(`已切换到 ${labelForMethod(targetMethod)}，正在刷新邮件`);
    await loadFolders({ silent: true });
    await loadMessages({ silent: true, keepSelection: false });
}

async function testMailboxConnection() {
    const payload = extractMailboxPayload();
    setBusy([elements.testMailboxButton], true);

    try {
        const response = await requestJson("/api/mailboxes/test-connection", {
            method: "POST",
            body: payload,
        });
        setStatus(response.message || `${labelForMethod(response.method)} 连接成功`);
    } catch (error) {
        setStatus(error.message, true);
    } finally {
        setBusy([elements.testMailboxButton], false);
    }
}

async function runBatchMessageAction(action, destinationFolder = "") {
    const messageIds = Array.from(state.selectedMessageIds);
    if (!messageIds.length) {
        setStatus("请先选择要处理的邮件", true);
        return;
    }

    const busyTargets = [
        elements.batchMarkReadButton,
        elements.batchMarkUnreadButton,
        elements.batchFlagButton,
        elements.batchUnflagButton,
        elements.batchMoveButton,
        elements.batchArchiveButton,
        elements.batchDeleteButton,
    ];
    setBusy(busyTargets, true);
    try {
        const payload = {
            mailbox_id: state.selectedMailboxId,
            method: state.activeMethod || state.selectedMailbox?.preferred_method || "graph_api",
            folder: state.selectedFolder || "INBOX",
            message_ids: messageIds,
            action,
        };
        if (destinationFolder) {
            payload.destination_folder = destinationFolder;
        }
        const response = await requestJson("/api/mailbox/messages/actions/batch", {
            method: "POST",
            body: payload,
        });
        state.selectedMessageIds = new Set();
        renderMessageBatchActions();
        await loadFolders({ silent: true });
        await loadMessages({ silent: true, keepSelection: false });
        const summary = response.summary || {};
        setStatus(`批量操作完成：成功 ${summary.succeeded || 0} 封，失败 ${summary.failed || 0} 封`);
    } catch (error) {
        setStatus(error.message, true);
    } finally {
        setBusy(busyTargets, false);
    }
}

async function selectAdjacentMessage(direction) {
    if (!state.messages.length) {
        return;
    }
    const currentIndex = state.messages.findIndex((item) => item.message_id === state.selectedMessageId);
    const nextIndex = currentIndex < 0 ? 0 : currentIndex + direction;
    if (nextIndex < 0 || nextIndex >= state.messages.length) {
        return;
    }
    await loadMessageDetail(state.messages[nextIndex].message_id);
}

function renderFolderList() {
    if (!elements.folderList) {
        return;
    }
    elements.folderList.innerHTML = "";
    if (!state.folders.length) {
        elements.folderList.innerHTML = "<div class='detail-empty compact-empty'>暂无文件夹数据</div>";
        setElementText(elements.folderSummary, "未加载");
        return;
    }
    const fragment = document.createDocumentFragment();
    state.folders.forEach((folder) => {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "folder-item";
        button.classList.toggle("is-active", folder.id === state.selectedFolder);
        button.innerHTML = `
            <strong>${escapeHtml(folder.display_name || folder.name || folder.id)}</strong>
            <span>${folder.unread || 0} 未读 / ${folder.total || 0} 总数</span>
        `;
        button.addEventListener("click", async () => {
            state.selectedFolder = folder.id;
            state.messagePage = 1;
            renderFolderList();
            await loadMessages({ silent: true, keepSelection: false });
        });
        fragment.append(button);
    });
    elements.folderList.replaceChildren(fragment);
    setElementText(elements.folderSummary, `共 ${state.folders.length} 个文件夹`);
}

function syncFolderSelectOptions() {
    [elements.moveFolderSelect, elements.batchMoveFolderSelect].forEach((select) => {
        if (!select) {
            return;
        }
        const currentValue = select.value;
        select.innerHTML = "<option value=''>移动到...</option>";
        state.folders.forEach((folder) => {
            const option = document.createElement("option");
            option.value = folder.id;
            option.textContent = folder.display_name || folder.name || folder.id;
            select.append(option);
        });
        if (currentValue) {
            select.value = currentValue;
        }
    });
}

function renderMessageBatchActions() {
    const selectedCount = state.selectedMessageIds.size;
    if (elements.messageBatchPanel) {
        elements.messageBatchPanel.hidden = selectedCount === 0;
    }
    setElementText(elements.selectedMessageCount, `已选 ${selectedCount} 封邮件`);
    [
        elements.clearSelectedMessagesButton,
        elements.batchMarkReadButton,
        elements.batchMarkUnreadButton,
        elements.batchFlagButton,
        elements.batchUnflagButton,
        elements.batchMoveButton,
        elements.batchArchiveButton,
        elements.batchDeleteButton,
    ].forEach((element) => {
        if (element) {
            element.disabled = selectedCount === 0;
        }
    });
    if (elements.batchMoveFolderSelect) {
        elements.batchMoveFolderSelect.disabled = !state.folders.length || selectedCount === 0;
    }
}

function renderMessagePagination() {
    const meta = state.messageMeta || createDefaultMessageMeta();
    const totalPages = meta.total_pages || 0;
    setElementText(elements.messagePageInfo, totalPages ? `第 ${meta.page} / ${totalPages} 页` : "第 0 / 0 页");
    elements.messagePrevPageButton.disabled = !meta.has_prev;
    elements.messageNextPageButton.disabled = !meta.has_next;
}

function renderMessages() {
    if (!state.messages.length) {
        const emptyState = document.createElement("div");
        emptyState.className = "detail-empty compact-empty";
        emptyState.textContent = "没有可显示的邮件";
        elements.messageList.replaceChildren(emptyState);
        return;
    }

    const listFragment = document.createDocumentFragment();
    state.messages.forEach((message) => {
        const itemFragment = elements.messageItemTemplate.content.cloneNode(true);
        const article = itemFragment.querySelector(".message-item-card");
        const checkbox = itemFragment.querySelector(".message-select-checkbox");
        const button = itemFragment.querySelector(".message-item");
        button.dataset.messageId = message.message_id;
        button.classList.toggle("is-active", message.message_id === state.selectedMessageId);
        button.classList.toggle("is-unread", !message.is_read);
        article.classList.toggle("is-selected", state.selectedMessageIds.has(message.message_id));
        checkbox.checked = state.selectedMessageIds.has(message.message_id);
        checkbox.addEventListener("change", () => {
            const nextSelected = new Set(state.selectedMessageIds);
            if (checkbox.checked) {
                nextSelected.add(message.message_id);
            } else {
                nextSelected.delete(message.message_id);
            }
            state.selectedMessageIds = nextSelected;
            renderMessageBatchActions();
            renderMessages();
        });
        button.querySelector(".message-item-sender").textContent = message.sender_name || message.sender || "未知发件人";
        button.querySelector(".message-item-time").textContent = formatDate(message.received_at);
        button.querySelector(".message-item-subject").textContent = message.subject || "无主题";
        button.querySelector(".message-item-preview").textContent = message.preview || "无正文预览";
        const flagBadge = button.querySelector(".message-item-flag");
        const importanceBadge = button.querySelector(".message-item-importance");
        flagBadge.hidden = !message.is_flagged;
        importanceBadge.hidden = !message.importance || message.importance === "normal";
        importanceBadge.textContent = importanceLabel(message.importance);
        button.addEventListener("click", async () => {
            await loadMessageDetail(message.message_id);
        });
        listFragment.append(itemFragment);
    });
    elements.messageList.replaceChildren(listFragment);
}

function renderDetail(message) {
    if (!message) {
        elements.detailEmpty.hidden = false;
        elements.detailCard.hidden = true;
        elements.markReadButton.disabled = true;
        elements.markUnreadButton.disabled = true;
        elements.flagMessageButton.disabled = true;
        elements.unflagMessageButton.disabled = true;
        elements.moveMessageButton.disabled = true;
        elements.archiveMessageButton.disabled = true;
        elements.deleteMessageButton.disabled = true;
        elements.prevMessageButton.disabled = true;
        elements.nextMessageButton.disabled = true;
        return;
    }

    elements.detailEmpty.hidden = true;
    elements.detailCard.hidden = false;
    setElementText(elements.detailMethodTag, labelForMethod(message.method));
    setElementText(elements.detailSubject, message.subject || "无主题");
    setElementText(elements.detailSender, formatSender(message));
    setElementText(elements.detailTime, formatDate(message.received_at, true));
    setElementText(elements.detailTo, formatRecipients(message.to_recipients || []));
    setElementText(elements.detailCc, formatRecipients(message.cc_recipients || []));
    setElementText(elements.detailBcc, formatRecipients(message.bcc_recipients || []));
    setElementText(elements.detailMessageId, message.internet_message_id || message.message_id || "-");
    setElementText(elements.detailConversationId, message.conversation_id || "-");
    setElementText(elements.detailReadBadge, message.is_read ? "已读" : "未读");
    setElementText(elements.detailFlagBadge, message.is_flagged ? "已星标" : "未星标");
    setElementText(elements.detailImportanceBadge, importanceLabel(message.importance));
    setElementText(elements.detailAttachmentBadge, message.has_attachments ? "有附件" : "无附件");
    setElementText(elements.detailBody, message.body_text || message.preview || "");
    setElementText(elements.detailHeaders, formatHeaders(message.headers || {}));
    renderAttachmentList(message.attachments || []);
    renderDetailView(message);
    elements.markReadButton.disabled = message.is_read;
    elements.markUnreadButton.disabled = !message.is_read;
    elements.flagMessageButton.disabled = message.is_flagged;
    elements.unflagMessageButton.disabled = !message.is_flagged;
    elements.moveMessageButton.disabled = false;
    elements.archiveMessageButton.disabled = false;
    elements.deleteMessageButton.disabled = false;
    elements.prevMessageButton.disabled = state.messages.findIndex((item) => item.message_id === message.message_id) <= 0;
    elements.nextMessageButton.disabled = state.messages.findIndex((item) => item.message_id === message.message_id) === state.messages.length - 1;
}

function fillMailboxForm(mailbox) {
    elements.mailboxForm.elements.mailbox_id.value = mailbox.id;
    elements.mailboxForm.elements.label.value = mailbox.label || "";
    elements.mailboxForm.elements.email.value = mailbox.email || "";
    elements.mailboxForm.elements.client_id.value = mailbox.client_id || "";
    elements.mailboxForm.elements.refresh_token.value = mailbox.refresh_token || "";
    elements.mailboxForm.elements.proxy.value = mailbox.proxy || "";
    elements.mailboxForm.elements.preferred_method.value = mailbox.preferred_method || "graph_api";
    elements.mailboxForm.elements.notes.value = mailbox.notes || "";
}

function resetMailboxForm() {
    elements.mailboxForm.reset();
    elements.mailboxForm.elements.mailbox_id.value = "";
    elements.mailboxForm.elements.preferred_method.value = "graph_api";
    elements.deleteMailboxButton.disabled = true;
}

function resetImportForm() {
    elements.importForm.reset();
    elements.importForm.elements.preferred_method.value = "graph_api";
    elements.importFileInput.value = "";
    setImportFileLabel("");
}

function resetWorkspace({ preserveAuth = false } = {}) {
    state.mailboxes = [];
    state.mailboxMeta = createDefaultMailboxMeta();
    state.mailboxSearchTerm = "";
    state.mailboxPage = 1;
    state.mailboxConnectionResults = {};
    state.selectedMailboxIds = new Set();
    state.bulkPreferredMethod = "";
    state.selectedMailboxId = null;
    state.selectedMailbox = null;
    state.selectedMailboxDetail = null;
    state.activeMethod = null;
    state.folders = [];
    state.selectedFolder = "INBOX";
    state.overviewMethods = [];
    state.messages = [];
    state.messageMeta = createDefaultMessageMeta();
    state.messagePage = 1;
    state.messagePageSize = 20;
    state.readState = "all";
    state.hasAttachmentsOnly = false;
    state.flaggedOnly = false;
    state.importance = "all";
    state.sortOrder = "desc";
    state.selectedMessageIds = new Set();
    state.selectedMessageId = null;
    state.selectedMessage = null;
    state.detailView = "text";
    state.searchTerm = "";
    state.mailboxFetchToken += 1;
    state.messageFetchToken += 1;
    state.detailFetchToken += 1;
    window.clearTimeout(state.mailboxSearchTimer);
    state.mailboxSearchTimer = null;
    stopAutoRefresh();
    elements.mailboxSearchInput.value = "";
    elements.bulkPreferredMethodSelect.value = "";
    elements.searchInput.value = "";
    elements.readStateSelect.value = "all";
    elements.importanceSelect.value = "all";
    elements.sortOrderSelect.value = "desc";
    elements.pageSizeSelect.value = "20";
    elements.attachmentOnlyCheckbox.checked = false;
    elements.flaggedOnlyCheckbox.checked = false;
    elements.lastSyncLabel.textContent = "未同步";
    renderMailboxList();
    renderOverview([]);
    renderFolderList();
    renderMessageBatchActions();
    renderMessagePagination();
    renderMessages();
    renderDetail(null);
    resetMailboxForm();
    resetImportForm();
    closeEditor();
    closeImportDrawer();
    clearPersistedWorkspaceState();
    if (!preserveAuth) {
        elements.currentMailboxLabel.textContent = "未选择";
    }
}

function extractMailboxPayload() {
    const formData = new FormData(elements.mailboxForm);
    return {
        label: readString(formData.get("label")),
        email: readString(formData.get("email")),
        client_id: readString(formData.get("client_id")),
        refresh_token: readString(formData.get("refresh_token")),
        proxy: readString(formData.get("proxy")),
        preferred_method: readString(formData.get("preferred_method")) || "graph_api",
        notes: readString(formData.get("notes")),
    };
}

function setMailboxSelection(mailboxId, checked) {
    const nextSelected = new Set(state.selectedMailboxIds);
    if (checked) {
        nextSelected.add(mailboxId);
    } else {
        nextSelected.delete(mailboxId);
    }
    state.selectedMailboxIds = nextSelected;
    persistWorkspaceState();
    renderMailboxList();
}

function getSelectedMailboxIds() {
    return Array.from(state.selectedMailboxIds);
}

function getVisibleSelectionState() {
    const visibleIds = state.mailboxes.map((item) => item.id);
    const selectedCountOnPage = visibleIds.filter((id) => state.selectedMailboxIds.has(id)).length;
    const allVisibleSelected = Boolean(visibleIds.length) && selectedCountOnPage === visibleIds.length;
    return {
        visibleCount: visibleIds.length,
        selectedCountOnPage,
        allVisibleSelected,
    };
}

function toggleVisibleMailboxesSelection() {
    if (!state.mailboxes.length) {
        return;
    }

    const { allVisibleSelected } = getVisibleSelectionState();
    const nextSelected = new Set(state.selectedMailboxIds);
    state.mailboxes.forEach((mailbox) => {
        if (allVisibleSelected) {
            nextSelected.delete(mailbox.id);
        } else {
            nextSelected.add(mailbox.id);
        }
    });
    state.selectedMailboxIds = nextSelected;
    persistWorkspaceState();
    renderMailboxList();
}

function clearSelectedMailboxes() {
    if (!state.selectedMailboxIds.size) {
        return;
    }
    state.selectedMailboxIds = new Set();
    persistWorkspaceState();
    renderMailboxList();
    setStatus("已清空跨页选择");
}

function renderBulkActions() {
    const totalSelected = state.selectedMailboxIds.size;
    const { visibleCount, selectedCountOnPage, allVisibleSelected } = getVisibleSelectionState();
    const pageHint = visibleCount ? `，本页 ${selectedCountOnPage} 个` : "";
    setElementText(elements.selectedMailboxCount, `已选 ${totalSelected} 个${pageHint}`);
    elements.toggleVisibleMailboxesButton.disabled = !visibleCount;
    elements.toggleVisibleMailboxesButton.textContent = allVisibleSelected ? "取消本页" : "本页全选";
    elements.clearSelectedMailboxesButton.disabled = totalSelected === 0;
    elements.testVisibleMailboxesButton.disabled = !visibleCount;
    elements.testSelectedMailboxesButton.disabled = totalSelected === 0;
    elements.bulkPreferredMethodSelect.value = state.bulkPreferredMethod;
    elements.applyPreferredMethodButton.disabled = totalSelected === 0 || !state.bulkPreferredMethod;
    elements.deleteSelectedMailboxesButton.disabled = totalSelected === 0;
}

function renderMailboxPagination({ selectedMailboxVisible }) {
    const meta = state.mailboxMeta || createDefaultMailboxMeta();
    const hasResults = meta.total > 0;
    const pageLabel = hasResults ? `第 ${meta.page} / ${meta.total_pages} 页` : "第 0 / 0 页";
    let resultLabel = state.mailboxSearchTerm ? `匹配 ${meta.total} 个邮箱` : `共 ${meta.total} 个邮箱`;

    if (state.selectedMailboxId && !selectedMailboxVisible) {
        resultLabel += "，当前选中邮箱未在本页";
    }

    setElementText(elements.mailboxResultInfo, resultLabel);
    setElementText(elements.mailboxPageInfo, pageLabel);
    elements.mailboxPrevPageButton.disabled = !meta.has_prev;
    elements.mailboxNextPageButton.disabled = !meta.has_next;
}

function revealMailboxInSidebar(mailboxId) {
    if (!mailboxId) {
        state.mailboxPage = 1;
        persistWorkspaceState();
        return;
    }

    state.mailboxSearchTerm = "";
    state.mailboxPage = 1;
    elements.mailboxSearchInput.value = "";
    persistWorkspaceState();
}

function createDefaultMailboxMeta() {
    return {
        q: "",
        page: 1,
        page_size: DEFAULT_MAILBOX_PAGE_SIZE,
        total: 0,
        total_pages: 0,
        has_prev: false,
        has_next: false,
    };
}

function createDefaultMessageMeta() {
    return {
        total: 0,
        returned: 0,
        page: 1,
        page_size: 20,
        total_pages: 0,
        has_prev: false,
        has_next: false,
        folder: "INBOX",
    };
}

function hydrateWorkspaceState() {
    try {
        const raw = window.sessionStorage.getItem(WORKSPACE_STORAGE_KEY);
        if (!raw) {
            return;
        }

        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === "object") {
            state.mailboxSearchTerm = typeof parsed.mailboxSearchTerm === "string" ? parsed.mailboxSearchTerm : "";
            state.mailboxPage = Number.isInteger(parsed.mailboxPage) && parsed.mailboxPage > 0 ? parsed.mailboxPage : 1;
            state.selectedMailboxIds = new Set(
                Array.isArray(parsed.selectedMailboxIds)
                    ? parsed.selectedMailboxIds.filter((item) => Number.isInteger(item))
                    : []
            );
            state.bulkPreferredMethod = typeof parsed.bulkPreferredMethod === "string" ? parsed.bulkPreferredMethod : "";
        }
    } catch (error) {
        console.warn("hydrateWorkspaceState failed", error);
        clearPersistedWorkspaceState();
    }

    elements.mailboxSearchInput.value = state.mailboxSearchTerm;
    elements.bulkPreferredMethodSelect.value = state.bulkPreferredMethod;
}

function persistWorkspaceState() {
    try {
        window.sessionStorage.setItem(
            WORKSPACE_STORAGE_KEY,
            JSON.stringify({
                mailboxSearchTerm: state.mailboxSearchTerm,
                mailboxPage: state.mailboxPage,
                selectedMailboxIds: getSelectedMailboxIds(),
                bulkPreferredMethod: state.bulkPreferredMethod,
            })
        );
    } catch (error) {
        console.warn("persistWorkspaceState failed", error);
    }
}

function clearPersistedWorkspaceState() {
    try {
        window.sessionStorage.removeItem(WORKSPACE_STORAGE_KEY);
    } catch (error) {
        console.warn("clearPersistedWorkspaceState failed", error);
    }
}

function ensureMailboxSelected() {
    if (!state.selectedMailboxId) {
        setStatus("请先选择一个邮箱档案", true);
        return false;
    }
    return true;
}

async function requestJson(url, options = {}) {
    const response = await fetch(url, {
        method: options.method || "GET",
        headers: {
            "Content-Type": "application/json",
        },
        body: options.body ? JSON.stringify(options.body) : undefined,
    });

    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
        if (response.status === 401) {
            applyAuthState({ authenticated: false });
        }
        throw new Error(payload?.error?.message || "请求失败");
    }
    return payload;
}

function setBusy(targets, busy) {
    targets.forEach((target) => {
        if (!target) {
            return;
        }
        target.classList.toggle("is-loading", busy);
        if (
            target instanceof HTMLButtonElement
            || target instanceof HTMLSelectElement
            || target instanceof HTMLInputElement
            || target instanceof HTMLTextAreaElement
        ) {
            target.disabled = busy;
        }
    });
}

function setStatus(message, isError = false) {
    elements.statusText.textContent = message;
    elements.statusText.style.color = isError ? "var(--red)" : "var(--ink)";
}

function labelForMethod(method) {
    return {
        graph_api: "Graph API",
        imap_new: "新版 IMAP",
        imap_old: "旧版 IMAP",
    }[method] || method;
}

function readString(value) {
    return typeof value === "string" ? value.trim() : "";
}

function setElementText(element, value) {
    if (element) {
        element.textContent = value;
    }
}

function findMessageById(messageId) {
    return state.messages.find((item) => item.message_id === messageId) || null;
}

function updateMessageSelection() {
    elements.messageList.querySelectorAll(".message-item").forEach((button) => {
        button.classList.toggle("is-active", button.dataset.messageId === state.selectedMessageId);
    });
}

function shouldRefetchDetail(summary) {
    if (!summary) {
        return false;
    }
    if (!state.selectedMessage) {
        return true;
    }
    if (state.selectedMessage.message_id !== summary.message_id) {
        return true;
    }
    return !Object.prototype.hasOwnProperty.call(state.selectedMessage, "body_text");
}

function mergeMessageSummary(detail, summary) {
    if (!detail) {
        return summary || null;
    }
    if (!summary) {
        return detail;
    }
    return { ...detail, ...summary };
}

function buildMessageListSignature(messages) {
    return messages
        .map((message) => [
            message.message_id || "",
            message.is_read ? "1" : "0",
            message.received_at || "",
            message.subject || "",
            message.sender || "",
            message.preview || "",
            message.has_attachments ? "1" : "0",
            message.internet_message_id || "",
            message.is_flagged ? "1" : "0",
            message.importance || "",
        ].join("::"))
        .join("|");
}

function formatImportSummary(summary) {
    if (!summary || typeof summary !== "object") {
        return "批量导入完成";
    }

    const processed = Number.isInteger(summary.processed) ? summary.processed : 0;
    const created = Number.isInteger(summary.created) ? summary.created : 0;
    const updated = Number.isInteger(summary.updated) ? summary.updated : 0;
    const deduplicated = Number.isInteger(summary.deduplicated) ? summary.deduplicated : 0;
    return `批量导入完成：处理 ${processed} 条，新建 ${created} 条，更新 ${updated} 条，去重 ${deduplicated} 条`;
}

function formatDate(value, withTime = false) {
    if (!value) {
        return "未知时间";
    }
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
        return value;
    }
    return new Intl.DateTimeFormat("zh-CN", {
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: withTime ? "2-digit" : undefined,
        minute: withTime ? "2-digit" : undefined,
    }).format(date);
}

function formatSender(message) {
    if (message.sender_name && message.sender) {
        return `${message.sender_name} <${message.sender}>`;
    }
    return message.sender || message.sender_name || "-";
}

function formatRecipients(values) {
    if (!values.length) {
        return "-";
    }
    return values.join("，");
}

function setImportFileLabel(fileName) {
    elements.importFileLabel.textContent = fileName || "支持 .txt / .csv / .tsv / .json";
}

function renderAttachmentList(attachments) {
    elements.detailAttachmentsSection.hidden = !attachments.length;
    elements.detailAttachmentList.innerHTML = "";
    if (!attachments.length) {
        return;
    }
    const fragment = document.createDocumentFragment();
    attachments.forEach((attachment) => {
        const item = document.createElement("article");
        item.className = "attachment-item";
        item.innerHTML = `
            <strong>${escapeHtml(attachment.name || "未命名附件")}</strong>
            <span>${escapeHtml(attachment.content_type || "unknown")} · ${formatBytes(attachment.size || 0)}</span>
        `;
        fragment.append(item);
    });
    elements.detailAttachmentList.replaceChildren(fragment);
}

function renderDetailView(message) {
    const canShowHtml = Boolean(message?.body_html);
    elements.detailHtmlViewButton.disabled = !canShowHtml;
    if (state.detailView === "html" && !canShowHtml) {
        state.detailView = "text";
    }

    elements.detailBody.hidden = state.detailView !== "text";
    elements.detailHtmlFrame.hidden = state.detailView !== "html";
    elements.detailHeaders.hidden = state.detailView !== "headers";

    elements.detailTextViewButton.classList.toggle("is-active", state.detailView === "text");
    elements.detailHtmlViewButton.classList.toggle("is-active", state.detailView === "html");
    elements.detailHeadersViewButton.classList.toggle("is-active", state.detailView === "headers");

    if (state.detailView === "html") {
        elements.detailHtmlFrame.srcdoc = message?.body_html || "<p>当前邮件没有 HTML 正文</p>";
    }
}

function formatHeaders(headers) {
    const entries = Object.entries(headers || {});
    if (!entries.length) {
        return "暂无 headers";
    }
    return entries.map(([key, value]) => `${key}: ${value}`).join("\n");
}

function importanceLabel(value) {
    return {
        high: "高优先级",
        low: "低优先级",
        normal: "普通",
        all: "全部",
    }[value] || "普通";
}

function formatBytes(value) {
    const size = Number(value) || 0;
    if (size < 1024) {
        return `${size} B`;
    }
    if (size < 1024 * 1024) {
        return `${(size / 1024).toFixed(1)} KB`;
    }
    return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

function escapeHtml(value) {
    return String(value)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll("\"", "&quot;")
        .replaceAll("'", "&#39;");
}





