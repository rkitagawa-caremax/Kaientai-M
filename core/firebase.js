(function () {
    'use strict';

    const firebaseConfig = {
        apiKey: "AIzaSyDmGbdR-fKINShSd9EneJrUMQCkdO8UFGs",
        authDomain: "kaientai-m-defence.firebaseapp.com",
        projectId: "kaientai-m-defence",
        storageBucket: "kaientai-m-defence.firebasestorage.app",
        messagingSenderId: "280719402775",
        appId: "1:280719402775:web:ebfdaaaa0ace09f6457952"
    };

    const ROOT_COLLECTION = 'kaientai_m_states';
    const CHUNK_COLLECTION = 'chunks';
    const CHUNK_SIZE = 700000;

    let app = null;
    let db = null;
    let initError = null;

    function padChunkId(index) {
        return String(index).padStart(4, '0');
    }

    function splitToChunks(text, size) {
        const chunks = [];
        for (let i = 0; i < text.length; i += size) {
            chunks.push(text.slice(i, i + size));
        }
        return chunks.length > 0 ? chunks : [''];
    }

    function ensureReady() {
        if (initError) throw initError;
        if (!db) throw new Error('Firebase is not initialized');
    }

    try {
        if (!window.firebase || !window.firebase.initializeApp) {
            throw new Error('Firebase SDK is not loaded');
        }
        app = window.firebase.initializeApp(firebaseConfig);
        db = window.firebase.firestore(app);
    } catch (err) {
        initError = err;
        console.warn('[KaientaiCloud] Firebase init failed:', err);
    }

    async function saveModuleState(moduleId, payload) {
        ensureReady();
        const moduleRef = db.collection(ROOT_COLLECTION).doc(moduleId);
        const serialized = JSON.stringify(payload);
        const chunks = splitToChunks(serialized, CHUNK_SIZE);

        let prevCount = 0;
        const prevMetaSnap = await moduleRef.get();
        if (prevMetaSnap.exists) {
            const prevMeta = prevMetaSnap.data() || {};
            prevCount = Number(prevMeta.chunkCount || 0);
        }

        await moduleRef.set({
            moduleId,
            chunkCount: chunks.length,
            byteLength: serialized.length,
            savedAt: window.firebase.firestore.FieldValue.serverTimestamp(),
            schemaVersion: 1
        }, { merge: true });

        const writes = [];
        for (let i = 0; i < chunks.length; i++) {
            writes.push(
                moduleRef.collection(CHUNK_COLLECTION).doc(padChunkId(i)).set({
                    index: i,
                    data: chunks[i]
                })
            );
        }

        for (let i = chunks.length; i < prevCount; i++) {
            writes.push(moduleRef.collection(CHUNK_COLLECTION).doc(padChunkId(i)).delete());
        }

        const chunked = 40;
        for (let i = 0; i < writes.length; i += chunked) {
            await Promise.all(writes.slice(i, i + chunked));
        }

        return { chunkCount: chunks.length, byteLength: serialized.length };
    }

    async function loadModuleState(moduleId) {
        ensureReady();
        const moduleRef = db.collection(ROOT_COLLECTION).doc(moduleId);
        const metaSnap = await moduleRef.get();
        if (!metaSnap.exists) return null;

        const meta = metaSnap.data() || {};
        const chunkCount = Number(meta.chunkCount || 0);
        if (chunkCount <= 0) return null;

        const reads = [];
        for (let i = 0; i < chunkCount; i++) {
            reads.push(moduleRef.collection(CHUNK_COLLECTION).doc(padChunkId(i)).get());
        }
        const snaps = await Promise.all(reads);
        const serialized = snaps.map(s => (s.exists ? (s.data() || {}).data || '' : '')).join('');
        if (!serialized) return null;
        return JSON.parse(serialized);
    }

    window.KaientaiCloud = {
        isReady() {
            return !!db && !initError;
        },
        getStatus() {
            return {
                ready: !!db && !initError,
                projectId: firebaseConfig.projectId,
                error: initError ? String(initError.message || initError) : ''
            };
        },
        saveModuleState,
        loadModuleState
    };
})();
