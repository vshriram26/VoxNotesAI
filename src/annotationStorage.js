const DATABASE_NAME = 'voxnotes-ai';
const DATABASE_VERSION = 1;
const STORE_NAME = 'annotation-sessions';

let databasePromise;

const isIndexedDbAvailable = () => {
  return typeof window !== 'undefined' && typeof window.indexedDB !== 'undefined';
};

const openDatabase = async () => {
  if (!isIndexedDbAvailable()) {
    return null;
  }

  if (!databasePromise) {
    databasePromise = new Promise((resolve, reject) => {
      const request = window.indexedDB.open(DATABASE_NAME, DATABASE_VERSION);

      request.onupgradeneeded = () => {
        const database = request.result;

        if (!database.objectStoreNames.contains(STORE_NAME)) {
          const store = database.createObjectStore(STORE_NAME, { keyPath: 'documentId' });
          store.createIndex('updatedAt', 'updatedAt', { unique: false });
        }
      };

      request.onsuccess = () => {
        resolve(request.result);
      };

      request.onerror = () => {
        reject(request.error || new Error('Unable to open the annotation database.'));
      };
    }).catch((error) => {
      databasePromise = undefined;
      throw error;
    });
  }

  return databasePromise;
};

export const getStoredAnnotations = async (documentId) => {
  if (!documentId) {
    return null;
  }

  const database = await openDatabase();

  if (!database) {
    return null;
  }

  return new Promise((resolve, reject) => {
    const transaction = database.transaction(STORE_NAME, 'readonly');
    const request = transaction.objectStore(STORE_NAME).get(documentId);

    request.onsuccess = () => {
      resolve(request.result ?? null);
    };

    request.onerror = () => {
      reject(request.error || new Error('Unable to read annotations from IndexedDB.'));
    };
  });
};

export const saveStoredAnnotations = async (record) => {
  if (!record?.documentId) {
    return;
  }

  const database = await openDatabase();

  if (!database) {
    throw new Error('IndexedDB is not available in this browser.');
  }

  return new Promise((resolve, reject) => {
    const transaction = database.transaction(STORE_NAME, 'readwrite');
    const request = transaction.objectStore(STORE_NAME).put(record);

    request.onsuccess = () => {
      resolve(record);
    };

    request.onerror = () => {
      reject(request.error || new Error('Unable to save annotations to IndexedDB.'));
    };
  });
};

export const listStoredAnnotations = async () => {
  const database = await openDatabase();

  if (!database) {
    return [];
  }

  return new Promise((resolve, reject) => {
    const transaction = database.transaction(STORE_NAME, 'readonly');
    const request = transaction.objectStore(STORE_NAME).getAll();

    request.onsuccess = () => {
      const records = Array.isArray(request.result) ? request.result : [];
      records.sort((leftRecord, rightRecord) => {
        return new Date(rightRecord.updatedAt || 0).getTime() - new Date(leftRecord.updatedAt || 0).getTime();
      });
      resolve(records);
    };

    request.onerror = () => {
      reject(request.error || new Error('Unable to list annotation sessions from IndexedDB.'));
    };
  });
};