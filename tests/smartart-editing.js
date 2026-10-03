/* Copyright (C) Ascensio System SIA, 2009-2026; SPDX-License-Identifier: AGPL-3.0-only */
(async function () {
    'use strict';
    const results = [];
    function assert(condition, message) { if (!condition) throw new Error(message); }
    function equal(actual, expected, message) {
        assert(JSON.stringify(actual) === JSON.stringify(expected), message + ': ' + JSON.stringify(actual));
    }
    try {
        assert(!testErrors.length, testErrors.join('\n'));
        AscCommon.History = AscCommon.History || new AscCommon.CHistory();
        window.History = AscCommon.History;
        AscCommon.CollaborativeEditing = new AscCommon.CCollaborativeEditingBase();
        const editorId = testProduct === 'cell' ? AscCommon.c_oEditorId.Spreadsheet :
            testProduct === 'slide' ? AscCommon.c_oEditorId.Presentation : AscCommon.c_oEditorId.Word;
        const prototype = testProduct === 'pdf' ? AscCommon.PDFEditorApi.prototype :
            testProduct === 'cell' ? Asc.spreadsheet_api.prototype : Asc.asc_docs_api.prototype;
        const worksheet = {getId: function () { return 'sheet'; }, getDrawingDocument: function () { return null; }};
        const pdfDoc = {Viewer: {getPageRotate: function () { return 0; }}, styles: {}};
        if (testProduct === 'pdf') AscPDF.CPDFDoc.prototype.InitDefaultTextListStyles.call(pdfDoc);
        Object.assign(pdfDoc, {AddToRedraw: function () {}, SetNeedUpdateSearch: function () {}, SetNeedUpdateTarget: function () {}, GetDrawingDocument: function () { return null; }});
        const api = window.editor = Asc.editor = Object.create(prototype);
        Object.assign(api, {
            isDocumentEditor: false, isViewMode: false, restrictions: Asc.c_oAscRestrictionType.None,
            WordControl: {m_oDrawingDocument: null, m_oLogicDocument: null},
            wbModel: {theme: AscFormat.GetDefaultTheme(), getWorksheetById: function (id) { return id === 'sheet' ? worksheet : null; }},
            frameManager: {getMainDiagramController: function () { return null; }},
            getEditorId: function () { return editorId; },
            getDrawingDocument: function () { return null; }, getLogicDocument: function () { return null; },
            getFontManager: function () { return null; }, asc_enableKeyEvents: function () {},
            isPdfEditor: function () { return testProduct === 'pdf'; },
            getDocumentRenderer: function () { return {paint: function () {}}; },
            getPDFDoc: function () { return pdfDoc; },
            checkObjectsLock: function (ids, callback) { setTimeout(function () { callback(!locked, true); }, 0); }
        });
        if (testProduct === 'slide') {
            const slide = {getObjectType: function () { return AscDFH.historyitem_type_Slide; },
                Layout: {Master: {Theme: AscFormat.GetDefaultTheme()}}, addToRecalculate: function () {}};
            api.WordControl.m_oLogicDocument = {GetCurrentSlide: function () { return slide; }, IsTrackRevisions: function () { return false; }, IsDocumentEditor: function () { return false; }, addToRecalculate: function () {}};
        }
        let locked = false;
        const changes = [];
        // Keep real history change objects, independent of document rendering.
        const history = AscCommon.History;
        history.Add = function (change) { if (history.TurnOffHistory === 0) changes.push(change); };
        history.CanAddChanges = function () { return history.TurnOffHistory === 0; };
        const controller = {
            selectedObjects: [], selection: {}, createShape: function () { return testProduct === 'pdf' ? new AscPDF.CPdfShape() : new AscFormat.CShape(); }, resetSelection: function () { this.selectedObjects = []; this.selection = {}; },
            selectObject: function (object) { this.selectedObjects = [object]; },
            checkSelectedObjectsAndCallback: function (callback) { if (!locked) callback(); }
        };
        api.getGraphicController = function () { return controller; };
        AscFonts.FontPickerByCharacter.checkText = function (text, editor, callback) { callback(); };
        AscCommon.g_oTableId.init();
        await AscCommon.g_oBinarySmartArts.checkLoadDrawing();
        let process;
        for (const [name, type] of Object.entries(Asc.c_oAscSmartArtTypes)) {
            try {
                await AscCommon.g_oBinarySmartArts.checkLoadData(type);
                const smartart = testProduct === 'pdf' ? new AscPDF.CPdfSmartArt() : new AscFormat.SmartArt();
                smartart.fillByPreset(type);
                if (!smartart.dataModel) continue;
                if (testProduct === 'cell') smartart.worksheet = worksheet;
                controller.selectObject(smartart);
                const before = api.asc_getSmartArtOutline();
                const oldDrawing = smartart.drawing;
                let edited;
                // Honor a preset's OOXML hierarchy and shape-count constraints.
                const last = before.nodes[before.nodes.length - 1];
                const attempts = [[before.nodes.length, last.depth], [before.nodes.length, last.depth + 1]];
                for (let i = 0; i <= before.nodes.length; ++i) {
                    for (let depth = 0; depth <= (i ? before.nodes[i - 1].depth + 1 : 0); ++depth) attempts.push([i, depth]);
                }
                for (const [position, depth] of attempts) {
                    const candidate = before.nodes.slice();
                    candidate.splice(position, 0, {text: 'New node: ä 中文', depth: depth});
                    if (smartart.prepareEditableOutline(candidate)) { edited = candidate; break; }
                }
                if (!edited) {
                    assert(!await api.asc_setSmartArtOutline(before.id, before.nodes.concat([{text: 'overflow', depth: 0}])), 'fixed layout accepted overflow');
                    for (let i = 0; i < before.nodes.length; ++i) {
                        const candidate = before.nodes.map(function (node, index) {
                            return Object.assign({}, node, {text: index === i ? 'New node: ä 中文' : node.text});
                        });
                        if (smartart.prepareEditableOutline(candidate)) { edited = candidate; break; }
                    }
                    assert(edited, 'preset has no editable text node');
                }
                changes.length = 0;
                assert(await api.asc_setSmartArtOutline(before.id, edited), 'apply failed');
                const after = smartart.getEditableOutline();
                const additions = after.filter(function (node) { return !before.nodes.some(function (item) { return item.id === node.id; }); });
                const pointMap = smartart.getDataModelFromData().getPtLst().getPtMap();
                after.filter(function (node) { return node.text === 'New node: ä 中文'; }).forEach(function (node) {
                    assert(pointMap[node.id].getPrSet().getPhldr() === false, 'edited text still marked as placeholder');
                    assert(typeof pointMap[node.id].getPhldrT() !== 'boolean', 'placeholder text contains boolean');
                });
                if (additions.length) {
                    equal(after.filter(function (node) { return before.nodes.some(function (item) { return item.id === node.id; }); }), before.nodes, 'existing nodes changed');
                    assert(additions.length === 1 && additions[0].text === 'New node: ä 中文', 'new text missing');
                } else assert(after.some(function (node) { return node.text === 'New node: ä 中文'; }), 'updated text missing');
                const newDrawing = smartart.drawing;
                assert(newDrawing.spTree.some(function (shape) {
                    const content = shape.getDocContent && shape.getDocContent();
                    return content && content.GetText({}).includes('New node: ä 中文');
                }), 'new node missing from rendered shapes');
                const action = changes.slice();
                history.TurnOff();
                for (let i = action.length - 1; i >= 0; --i) action[i].Undo();
                equal(smartart.getEditableOutline(), before.nodes, 'undo failed');
                assert(smartart.drawing === oldDrawing, 'undo drawing failed');
                for (const change of action) change.Redo();
                equal(smartart.getEditableOutline(), after, 'redo failed');
                assert(smartart.drawing === newDrawing, 'redo drawing failed');
                history.TurnOn();
                const writer = new AscCommon.CBinaryFileWriter();
                writer.StartRecord(0);
                smartart.toPPTY(writer);
                writer.EndRecord();
                const bytes = writer.GetData();
                const reader = new AscCommon.BinaryPPTYLoader();
                reader.stream = new AscCommon.FileStream(bytes, bytes.length);
                reader.stream.GetUChar();
                const reopened = testProduct === 'pdf' ? new AscPDF.CPdfSmartArt() : new AscFormat.SmartArt();
                reopened.fromPPTY(reader);
                equal(reopened.getEditableOutline(), after, 'saved outline changed');
                const reopenedPoints = reopened.getDataModelFromData().getPtLst().getPtMap();
                after.filter(function (node) { return node.text === 'New node: ä 中文'; }).forEach(function (node) {
                    assert(reopenedPoints[node.id].getPrSet().getPhldr() === false, 'saved text reopened as placeholder');
                });
                assert(reopened.drawing.spTree.length === newDrawing.spTree.length, 'saved drawing changed');
                assert(await api.asc_setSmartArtOutline(before.id, before.nodes), 'remove failed');
                equal(smartart.getEditableOutline(), before.nodes, 'remove changed nodes');
                if (name === 'BasicProcess') process = smartart;
                results.push({name: name, passed: true, operation: additions.length ? 'add/remove' : 'text/fixed-layout'});
            } catch (error) {
                // A failure must not leave the editor's history disabled.
                history.TurnOffHistory = 0;
                AscCommon.g_oTableId.m_bTurnOff = false;
                results.push({name: name, error: error.stack});
            }
        }
        assert(process, 'BasicProcess missing');
        if (AscCommon.CShapeDrawer && AscCommon.CGraphics) {
            const canvas = document.createElement('canvas');
            canvas.width = canvas.height = 100;
            const context = canvas.getContext('2d');
            const graphics = new AscCommon.CGraphics();
            graphics.init(context, 100, 100, 20, 20);
            graphics.transform(1, 0, 0, 1, 0, 0);
            const overlay = new AscFormat.OverlayObject(AscFormat.CreateGeometry('rect'), 5, 5,
                AscFormat.CreateSolidFillRGBA(30, 90, 200, 255), null, new AscCommon.CMatrix());
            assert(!overlay.shapeDrawer, 'preview renderer allocated during layout creation');
            overlay.updateExtents(5, 5);
            overlay.draw(graphics);
            assert(overlay.shapeDrawer, 'preview renderer missing on draw');
            assert(context.getImageData(10, 10, 1, 1).data[3] > 0, 'preview draw produced no pixels');
        }
        controller.selectObject(process);
        const snapshot = api.asc_getSmartArtOutline();
        const reordered = snapshot.nodes.slice().reverse();
        reordered[1] = Object.assign({}, reordered[1], {depth: 1, text: 'Child', assistant: true});
        assert(await api.asc_setSmartArtOutline(snapshot.id, reordered), 'reparent failed');
        equal(process.getEditableOutline().map(function (node) { return [node.id, node.depth, node.assistant]; }),
            reordered.map(function (node) { return [node.id, node.depth, node.assistant]; }), 'hierarchy failed');
        const unchanged = process.getEditableOutline();
        assert(!await api.asc_setSmartArtOutline('stale-selection', reordered), 'stale selection accepted');
        assert(!await api.asc_setSmartArtOutline(snapshot.id, []), 'empty outline accepted');
        assert(!await api.asc_setSmartArtOutline(snapshot.id, [{text: 'bad', depth: 1}]), 'bad hierarchy accepted');
        assert(!await api.asc_setSmartArtOutline(snapshot.id, [unchanged[0], unchanged[0]]), 'duplicate ids accepted');
        assert(!await api.asc_setSmartArtOutline(snapshot.id, [{id: null, text: 'bad', depth: 0}]), 'bad id accepted');
        assert(!await api.asc_setSmartArtOutline(snapshot.id, [{text: 'bad', depth: 0, assistant: 'yes'}]), 'bad assistant flag accepted');
        api.isViewMode = true;
        assert(!await api.asc_setSmartArtOutline(snapshot.id, unchanged), 'view mode accepted');
        api.isViewMode = false;
        locked = true;
        assert(!await api.asc_setSmartArtOutline(snapshot.id, unchanged), 'locked edit accepted');
        locked = false;
        equal(process.getEditableOutline(), unchanged, 'rejected edits changed data');
        controller.selection.groupSelection = {selectedObjects: [process.drawing.spTree[0]]};
        assert(api.asc_getSmartArtOutline().id === snapshot.id, 'internal shape selection failed');
        controller.selection = {};
        const group = new AscFormat.CGroupShape();
        group.addToSpTree(0, process);
        process.setGroup(group);
        group.selectedObjects = [process];
        controller.selectObject(group);
        controller.selection.groupSelection = group;
        assert(await api.asc_setSmartArtOutline(snapshot.id, unchanged), 'grouped edit failed');
        assert(controller.selectedObjects[0] === group && controller.selection.groupSelection === group, 'group selection lost');
        assert(api.asc_getSmartArtOutline().id === snapshot.id, 'grouped selection lost');
        process.setGroup(null);
        controller.selectObject(process);
        controller.selection = {};

        const checkFont = AscFonts.FontPickerByCharacter.checkText;
        let fontsReady;
        AscFonts.FontPickerByCharacter.checkText = function (text, editor, callback) { fontsReady = callback; };
        let pending = api.asc_setSmartArtOutline(snapshot.id, unchanged);
        controller.selectedObjects = [];
        fontsReady();
        assert(!await pending, 'selection changed during font loading');
        controller.selectObject(process);
        pending = api.asc_setSmartArtOutline(snapshot.id, unchanged);
        api.isViewMode = true;
        fontsReady();
        assert(!await pending, 'permissions changed during font loading');
        api.isViewMode = false;
        pending = api.asc_setSmartArtOutline(snapshot.id, unchanged);
        controller.checkSelectedObjectsProtection = function () { return true; };
        fontsReady();
        assert(!await pending, 'protection changed during font loading');
        delete controller.checkSelectedObjectsProtection;
        pending = api.asc_setSmartArtOutline(snapshot.id, unchanged);
        api.collaborativeEditing = {getGlobalLock: function () { return true; }};
        fontsReady();
        assert(!await pending, 'collaborative lock changed during font loading');
        delete api.collaborativeEditing;
        pending = api.asc_setSmartArtOutline(snapshot.id, unchanged);
        const concurrent = unchanged.map(function (node, index) {
            return Object.assign({}, node, {text: index === 0 ? 'Edited elsewhere' : node.text});
        });
        process.applyEditableOutline(process.prepareEditableOutline(concurrent, true));
        fontsReady();
        assert(!await pending, 'concurrent edit overwritten');
        equal(process.getEditableOutline(), concurrent, 'concurrent text lost');
        AscFonts.FontPickerByCharacter.checkText = checkFont;
        assert(await api.asc_setSmartArtOutline(snapshot.id, unchanged), 'restore after concurrent edit failed');
        equal(process.getEditableOutline(), unchanged, 'async rejection changed diagram');

        // Exercise each product's real lock/action adapter with a minimal host.
        const controllerPrototype = testProduct === 'pdf' ? AscPDF.CGraphicObjects.prototype :
            testProduct === 'word' ? AscCommonWord.CGraphicObjects.prototype : AscFormat.DrawingObjectsController.prototype;
        const adapterEvents = [];
        const logicDocument = api.WordControl.m_oLogicDocument;
        const host = logicDocument || {};
        Object.assign(host, {
            IsDocumentEditor: function () { return false; }, IsTrackRevisions: function () { return false; },
            Get_ColorMap: function () { return AscFormat.GetDefaultColorMap(); },
            Document_Is_SelectionLocked: function () { return locked; },
            StartAction: function () { adapterEvents.push('start'); },
            Recalculate: function () { adapterEvents.push('recalculate'); },
            FinalizeAction: function () { adapterEvents.push('finish'); }
        });
        api.WordControl.m_oLogicDocument = host;
        Object.assign(controller, {
            document: host, drawingObjects: {sendGraphicObjectProps: function () { adapterEvents.push('properties'); }},
            getPresentation: function () { return host; }, getEditorApi: function () { return api; },
            getSelectionState: function () { return []; },
            startRecalculate: function () { adapterEvents.push('recalculate'); },
            checkObjectsAndCallback: controllerPrototype.checkObjectsAndCallback,
            checkSelectedObjectsAndCallback: controllerPrototype.checkSelectedObjectsAndCallback
        });
        api.sendEvent = function (name) { adapterEvents.push(name); };
        api.checkChangesSize = function () {};
        const createPoint = history.Create_NewPoint;
        history.Create_NewPoint = function () { adapterEvents.push('start'); };
        const adapterEdit = unchanged.map(function (node, index) {
            return Object.assign({}, node, {text: index === 0 ? 'Real editor adapter' : node.text});
        });
        assert(await api.asc_setSmartArtOutline(snapshot.id, adapterEdit), 'real editor adapter rejected edit');
        assert(adapterEvents.includes('start'), 'real editor did not start an action');
        assert(adapterEvents.includes(testProduct === 'cell' ? 'asc_onUserActionEnd' : 'finish'), 'real editor did not finish the action');
        if (testProduct !== 'slide') assert(adapterEvents.includes('recalculate'), 'real editor did not recalculate');
        equal(process.getEditableOutline(), adapterEdit, 'real editor adapter lost text');
        adapterEvents.length = 0;
        locked = true;
        assert(!await api.asc_setSmartArtOutline(snapshot.id, unchanged), 'real editor adapter ignored lock');
        assert(!adapterEvents.includes('start'), 'locked editor started an action');
        equal(process.getEditableOutline(), adapterEdit, 'locked editor changed text');
        locked = false;
        assert(await api.asc_setSmartArtOutline(snapshot.id, unchanged), 'real editor adapter restore failed');
        history.Create_NewPoint = createPoint;
        api.WordControl.m_oLogicDocument = logicDocument;
        controller.checkSelectedObjectsAndCallback = function (callback) { if (!locked) callback(); };

        // Exercise the actual shared UI, including hierarchy and Cancel.
        window.OnlyOfficeSmartArtDialog(api);
        let dialog = document.querySelector('dialog');
        assert(dialog && dialog.open, 'dialog did not open');
        const click = function (caption) {
            const button = Array.from(dialog.querySelectorAll('button')).find(function (item) { return item.textContent === caption; });
            assert(button && !button.disabled, 'button disabled: ' + caption);
            button.click();
        };
        click('Add node');
        click('Demote');
        click('Promote');
        click('Move up');
        click('Move down');
        click('Remove branch');
        click('Cancel');
        await new Promise(function (resolve) { setTimeout(resolve, 50); });
        equal(process.getEditableOutline(), unchanged, 'cancel changed diagram');
        window.OnlyOfficeSmartArtDialog(api);
        dialog = document.querySelector('dialog');
        const input = dialog.querySelector('textarea');
        input.value = 'Edited from dialog\nSecond line';
        input.dispatchEvent(new Event('input'));
        click('Apply');
        await new Promise(function (resolve) { setTimeout(resolve, 50); });
        assert(process.getEditableOutline()[0].text === 'Edited from dialog\nSecond line', 'dialog apply failed');
        assert(!document.querySelector('dialog'), 'dialog did not close');
        // Older mobile browsers do not implement native dialog.showModal.
        const showModal = HTMLDialogElement.prototype.showModal;
        HTMLDialogElement.prototype.showModal = undefined;
        window.OnlyOfficeSmartArtDialog(api);
        dialog = document.querySelector('dialog');
        assert(dialog && dialog.hasAttribute('open') && dialog.parentElement !== document.body, 'fallback did not open');
        const lastButton = dialog.querySelectorAll('button');
        lastButton[lastButton.length - 1].focus();
        dialog.dispatchEvent(new KeyboardEvent('keydown', {key: 'Tab', bubbles: true}));
        assert(document.activeElement === lastButton[0], 'fallback focus trap failed');
        dialog.dispatchEvent(new KeyboardEvent('keydown', {key: 'Escape', bubbles: true}));
        assert(!document.querySelector('dialog'), 'fallback did not close');
        HTMLDialogElement.prototype.showModal = showModal;
        window.Common = window.Common || {};
        const namespace = testProduct === 'cell' ? 'SSE' : testProduct === 'slide' ? 'PE' : testProduct === 'pdf' ? 'PDFE' : 'DE';
        const app = testProduct === 'cell' ? 'spreadsheeteditor' : testProduct === 'slide' ? 'presentationeditor' : testProduct === 'pdf' ? 'pdfeditor' : 'documenteditor';
        window[namespace] = {Views: {}};
        await new Promise(function (resolve, reject) {
            const script = document.createElement('script');
            script.src = '../../web-apps/vendor/requirejs/require.js';
            script.onload = resolve;
            script.onerror = reject;
            document.head.appendChild(script);
        });
        require.config({
            baseUrl: '../../web-apps/apps',
            paths: {
                jquery: '../vendor/jquery/jquery', underscore: '../vendor/underscore/underscore',
                backbone: '../vendor/backbone/backbone', text: '../vendor/requirejs-text/text',
                jmousewheel: '../vendor/perfect-scrollbar/src/jquery.mousewheel',
                perfectscrollbar: 'common/main/lib/mods/perfect-scrollbar',
                core: 'common/main/lib/core/application', notification: 'common/main/lib/core/NotificationCenter',
                irregularstack: 'common/IrregularStack', gateway: 'common/Gateway', keymaster: 'common/main/lib/core/keymaster'
            },
            shim: {
                core: {deps: ['backbone', 'notification', 'irregularstack']}, notification: {deps: ['backbone']},
                backbone: {deps: ['jquery', 'underscore'], exports: 'Backbone'}, gateway: {deps: ['jquery']},
                perfectscrollbar: {deps: ['jmousewheel']}, jmousewheel: {deps: ['jquery']}
            }
        });
        const load = function (modules) {
            return new Promise(function (resolve, reject) {
                require(modules, function () { resolve(Array.from(arguments)); }, reject);
            });
        };
        const libraries = await load(['jquery', 'underscore', 'backbone']);
        window.$ = window.jQuery = libraries[0];
        window._ = libraries[1];
        window.Backbone = libraries[2];
        await load(['common/main/lib/mods/dropdown', 'common/main/lib/mods/tooltip', 'common/main/lib/util/utils',
            'notification', 'common/main/lib/util/LocalStorage', 'common/main/lib/controller/Scaling']);
        new Backbone.Application({nameSpace: namespace, autoCreate: false, controllers: []});
        await load(['common/main/lib/component/Window', 'common/main/lib/component/Menu', 'common/main/lib/component/Button']);
        await load([app + '/main/app/view/ShapeSettings']);
        const container = document.createElement('div');container.id='id-shape-settings';document.body.appendChild(container);
        const panel = new window[namespace].Views.ShapeSettings();
        panel.api = window.editor;
        if (!panel.btnEditSmartArt || !container.querySelector('.smartart-edit button')) throw Error('sidebar button missing');
        container.querySelector('.smartart-edit button').click();
        await new Promise(resolve=>setTimeout(resolve,10));
        if (!document.querySelector('dialog')) throw Error('sidebar button did not open dialog');
        assert(panel.lockedControls.includes(panel.btnEditSmartArt), 'sidebar button missing from locked controls');
        Array.from(document.querySelector('dialog').querySelectorAll('button')).find(function (button) { return button.textContent === 'Cancel'; }).click();
        assert(!document.querySelector('dialog'), 'sidebar dialog did not close');
        assert(!testErrors.length, testErrors.join('\n'));
        results.push({name: 'API, hierarchy, permissions, shared dialog and real sidebar', passed: true});
    } catch (error) {
        results.push({name: 'harness', error: error.stack});
    }
    document.getElementById('result').textContent = JSON.stringify({product: testProduct, results: results});
})();
