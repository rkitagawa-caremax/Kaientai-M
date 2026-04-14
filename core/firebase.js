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

    const AUTHORIZED_EMAILS = [
        'r.kitagawa@g.caremax.co.jp',
        'k.ogura@g.caremax.co.jp'
    ];
    const AUTHORIZED_EMAIL_SET = new Set(AUTHORIZED_EMAILS.map(email => normalizeEmail(email)));
    const ROOT_COLLECTION = 'kaientai_m_states';
    const CHUNK_COLLECTION = 'chunks';
    const CHUNK_SIZE = 120000;

    let app = null;
    let db = null;
    let auth = null;
    let initError = null;
    let authReady = false;
    let authPending = false;
    let currentUser = null;
    let currentAuthorized = false;
    let authMessage = '認証状態を確認しています...';
    const authListeners = new Set();

    function normalizeEmail(email) {
        return String(email || '').trim().toLowerCase();
    }

    function isAuthorizedEmail(email) {
        return AUTHORIZED_EMAIL_SET.has(normalizeEmail(email));
    }

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

    function ensureFirestoreReady() {
        if (initError) throw initError;
        if (!db) throw new Error('Firebase is not initialized');
    }

    function ensureCloudAccess() {
        ensureFirestoreReady();
        if (auth && !currentAuthorized) {
            throw new Error('Authorized Google sign-in is required');
        }
    }

    function ensureAuthReady() {
        if (initError) throw initError;
        if (!auth) throw new Error('Firebase Auth is not initialized');
    }

    function getProvider() {
        const provider = new window.firebase.auth.GoogleAuthProvider();
        provider.setCustomParameters({ prompt: 'select_account' });
        return provider;
    }

    function getAuthStatus() {
        return {
            ready: !!auth && !initError && authReady,
            pending: authPending,
            signedIn: !!currentUser,
            authorized: currentAuthorized,
            email: currentUser?.email || '',
            displayName: currentUser?.displayName || '',
            allowedEmails: AUTHORIZED_EMAILS.slice(),
            message: authMessage,
            error: initError ? String(initError.message || initError) : ''
        };
    }

    function syncAuthUi(status) {
        if (document.body) {
            document.body.classList.toggle('app-auth-locked', !status.authorized);
        }

        const allowedList = document.getElementById('auth-allowed-list');
        if (allowedList && allowedList.childElementCount === 0) {
            allowedList.textContent = '';
            status.allowedEmails.forEach(email => {
                const li = document.createElement('li');
                li.textContent = email;
                allowedList.appendChild(li);
            });
        }

        const signInBtn = document.getElementById('auth-signin-btn');
        if (signInBtn) {
            signInBtn.hidden = !!status.signedIn;
            signInBtn.disabled = status.pending || !!status.error;
            signInBtn.textContent = status.pending ? 'ログイン中...' : 'Googleでログイン';
        }

        const gateSignInBtn = document.getElementById('auth-gate-signin-btn');
        if (gateSignInBtn) {
            gateSignInBtn.disabled = status.pending || !!status.error;
            gateSignInBtn.textContent = status.pending ? 'ログイン中...' : 'Googleでログイン';
        }

        const signOutBtn = document.getElementById('auth-signout-btn');
        if (signOutBtn) {
            signOutBtn.hidden = !status.signedIn;
            signOutBtn.disabled = status.pending;
        }

        const userLabel = document.getElementById('auth-user-label');
        if (userLabel) {
            userLabel.textContent = status.authorized
                ? (status.email || '認証済み')
                : (status.signedIn && status.email ? status.email : '未ログイン');
        }

        const messageEl = document.getElementById('auth-gate-message');
        if (messageEl) {
            messageEl.textContent = status.error
                ? `認証初期化に失敗しました: ${status.error}`
                : (status.message || '承認済みの Google アカウントでログインしてください。');
        }
    }

    function notifyAuthListeners() {
        const status = getAuthStatus();
        syncAuthUi(status);
        authListeners.forEach(listener => {
            try {
                listener(status);
            } catch (err) {
                console.warn('[KaientaiAuth] listener failed:', err);
            }
        });
    }

    async function signInWithGoogle() {
        ensureAuthReady();
        authPending = true;
        authMessage = 'Googleログインを開始しています...';
        notifyAuthListeners();

        try {
            const result = await auth.signInWithPopup(getProvider());
            const email = normalizeEmail(result?.user?.email);
            if (email && !isAuthorizedEmail(email)) {
                authMessage = `このアカウントは許可されていません: ${email}`;
                notifyAuthListeners();
                try {
                    await auth.signOut();
                } catch (signOutErr) {
                    console.warn('[KaientaiAuth] signOut after unauthorized login failed:', signOutErr);
                }
                throw new Error(authMessage);
            }
            authMessage = email ? `ログイン済み: ${email}` : 'ログインに成功しました。';
            return result?.user || null;
        } catch (err) {
            if (err?.code === 'auth/popup-closed-by-user') {
                authMessage = 'ログインがキャンセルされました。';
            } else if (err?.code === 'auth/cancelled-popup-request') {
                authMessage = '別のログイン要求を処理中です。';
            } else if (!(err?.message || '').startsWith('このアカウントは許可されていません')) {
                authMessage = `Googleログインに失敗しました: ${err?.message || err}`;
            }
            throw err;
        } finally {
            authPending = false;
            notifyAuthListeners();
        }
    }

    async function signOut() {
        ensureAuthReady();
        authPending = true;
        authMessage = 'ログアウトしています...';
        notifyAuthListeners();

        try {
            await auth.signOut();
            authMessage = 'ログアウトしました。承認済みの Google アカウントでログインしてください。';
        } finally {
            authPending = false;
            notifyAuthListeners();
        }
    }

    function onAuthStateChanged(listener) {
        if (typeof listener !== 'function') return () => {};
        authListeners.add(listener);
        listener(getAuthStatus());
        return () => authListeners.delete(listener);
    }

    function bindAuthUi() {
        const signInBtn = document.getElementById('auth-signin-btn');
        if (signInBtn && signInBtn.dataset.bound !== '1') {
            signInBtn.dataset.bound = '1';
            signInBtn.addEventListener('click', () => {
                signInWithGoogle().catch(() => {});
            });
        }

        const gateSignInBtn = document.getElementById('auth-gate-signin-btn');
        if (gateSignInBtn && gateSignInBtn.dataset.bound !== '1') {
            gateSignInBtn.dataset.bound = '1';
            gateSignInBtn.addEventListener('click', () => {
                signInWithGoogle().catch(() => {});
            });
        }

        const signOutBtn = document.getElementById('auth-signout-btn');
        if (signOutBtn && signOutBtn.dataset.bound !== '1') {
            signOutBtn.dataset.bound = '1';
            signOutBtn.addEventListener('click', () => {
                signOut().catch(err => {
                    console.warn('[KaientaiAuth] signOut failed:', err);
                });
            });
        }

        notifyAuthListeners();
    }

    try {
        if (!window.firebase || !window.firebase.initializeApp) {
            throw new Error('Firebase SDK is not loaded');
        }

        app = window.firebase.apps && window.firebase.apps.length
            ? window.firebase.app()
            : window.firebase.initializeApp(firebaseConfig);
        db = window.firebase.firestore(app);

        if (!window.firebase.auth) {
            throw new Error('Firebase Auth SDK is not loaded');
        }

        auth = window.firebase.auth();
        auth.onAuthStateChanged(async (user) => {
            authReady = true;

            if (!user) {
                currentUser = null;
                currentAuthorized = false;
                if (
                    !authMessage ||
                    authMessage.startsWith('認証状態を確認') ||
                    authMessage.startsWith('Googleログインを開始') ||
                    authMessage.startsWith('ログイン済み')
                ) {
                    authMessage = '承認済みの Google アカウントでログインしてください。';
                }
                notifyAuthListeners();
                return;
            }

            const email = normalizeEmail(user.email);
            if (!isAuthorizedEmail(email)) {
                currentUser = null;
                currentAuthorized = false;
                authMessage = `このアカウントは許可されていません: ${email}`;
                notifyAuthListeners();
                try {
                    await auth.signOut();
                } catch (signOutErr) {
                    console.warn('[KaientaiAuth] signOut after auth state check failed:', signOutErr);
                }
                return;
            }

            currentUser = user;
            currentAuthorized = true;
            authMessage = `ログイン済み: ${email}`;
            notifyAuthListeners();
        });
    } catch (err) {
        initError = err;
        authReady = true;
        authMessage = `認証初期化に失敗しました: ${err.message || err}`;
        console.warn('[KaientaiCloud] Firebase init failed:', err);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bindAuthUi, { once: true });
    } else {
        bindAuthUi();
    }

    async function saveModuleState(moduleId, payload) {
        ensureCloudAccess();
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
        ensureCloudAccess();
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

    window.KaientaiAuth = {
        isReady() {
            return !!auth && !initError && authReady;
        },
        isAuthorized() {
            return currentAuthorized;
        },
        getStatus: getAuthStatus,
        signInWithGoogle,
        signOut,
        onAuthStateChanged,
        getAllowedEmails() {
            return AUTHORIZED_EMAILS.slice();
        }
    };

    window.KaientaiCloud = {
        isReady() {
            return !!db && !initError && (!auth || currentAuthorized);
        },
        getStatus() {
            return {
                ready: !!db && !initError,
                authReady,
                signedIn: !!currentUser,
                authorized: currentAuthorized,
                projectId: firebaseConfig.projectId,
                error: initError ? String(initError.message || initError) : ''
            };
        },
        saveModuleState,
        loadModuleState
    };
})();



