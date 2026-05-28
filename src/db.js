import Dexie from 'dexie';
import { firestore, isCloudActive } from './firebase';
import { 
  collection, 
  getDocs, 
  addDoc, 
  updateDoc, 
  deleteDoc, 
  query, 
  where, 
  doc, 
  setDoc 
} from 'firebase/firestore';

export const db = new Dexie('YTAuditDatabase');

// Cấu trúc bảng IndexedDB
db.version(2).stores({
  channels: '++id, name, link, assignedStaff, partner, startDate, createdAt',
  staff: '++id, name, email, createdAt',
  metrics: '++id, channelId, periodStart, periodEnd, isLifetime, createdAt',
  audits: '++id, channelId, staffId, metricsPeriodId, metricsBaselineId, createdAt'
});

// Phiên bản 3: Hỗ trợ phân quyền tài khoản (username, password, role)
db.version(3).stores({
  channels: '++id, name, link, assignedStaff, partner, startDate, createdAt',
  staff: '++id, name, email, username, password, role, createdAt',
  metrics: '++id, channelId, periodStart, periodEnd, isLifetime, createdAt',
  audits: '++id, channelId, staffId, metricsPeriodId, metricsBaselineId, createdAt'
});

// --- HELPER PHÁT SINH ID SỐ TƯƠNG THÍCH LỚN (Firestore & Local) ---
function generateNumericId() {
  return Date.now() + Math.floor(Math.random() * 1000);
}

// --- XÁC THỰC ĐĂNG NHẬP (CẬP NHẬT RÀNG BUỘC 1 THIẾT BỊ) ---
export async function validateLogin(username, password, currentDeviceId = '') {
  if (isCloudActive) {
    try {
      const q = query(
        collection(firestore, 'staff'),
        where('username', '==', username),
        where('password', '==', password)
      );
      const snap = await getDocs(q);
      if (!snap.empty) {
        const user = snap.docs[0].data();
        
        // Chỉ giới hạn 1 thiết bị duy nhất đối với vai trò nhân viên (employee)
        if (user.role === 'employee' && currentDeviceId) {
          if (user.deviceId && user.deviceId !== currentDeviceId) {
            return { 
              success: false, 
              error: '⚠️ CẢNH BÁO BẢO MẬT: Tài khoản này đã được liên kết với một thiết bị khác! Vui lòng liên hệ Admin hoặc Trưởng phòng để reset liên kết thiết bị mới.' 
            };
          }
          if (!user.deviceId) {
            // Liên kết thiết bị lần đầu tiên
            const docRef = snap.docs[0].ref;
            await updateDoc(docRef, { deviceId: currentDeviceId });
            user.deviceId = currentDeviceId;
          }
        }
        
        return { success: true, user };
      }
      return { success: false, error: 'Tài khoản hoặc mật khẩu không chính xác!' };
    } catch (err) {
      console.error('Error validating login in Firestore:', err);
      return { success: false, error: `Lỗi Firebase: ${err.message}` };
    }
  } else {
    try {
      const matched = await db.staff
        .where('username')
        .equals(username)
        .and(s => s.password === password)
        .first();
      if (matched) {
        // Chỉ giới hạn 1 thiết bị duy nhất đối với vai trò nhân viên (employee)
        if (matched.role === 'employee' && currentDeviceId) {
          if (matched.deviceId && matched.deviceId !== currentDeviceId) {
            return { 
              success: false, 
              error: '⚠️ CẢNH BÁO BẢO MẬT: Tài khoản này đã được liên kết với một thiết bị khác! Vui lòng liên hệ Admin hoặc Trưởng phòng để reset liên kết thiết bị mới.' 
            };
          }
          if (!matched.deviceId) {
            await db.staff.update(matched.id, { deviceId: currentDeviceId });
            matched.deviceId = currentDeviceId;
          }
        }
        return { success: true, user: matched };
      }
      return { success: false, error: 'Tài khoản hoặc mật khẩu không chính xác!' };
    } catch (err) {
      return { success: false, error: `Lỗi Local DB: ${err.message}` };
    }
  }
}

// --- LIÊN KẾT / RESET THIẾT BỊ DUY NHẤT CHO NHÂN VIÊN ---
export async function bindStaffDevice(staffId, deviceId) {
  const numId = Number(staffId);
  if (isCloudActive) {
    const q = query(collection(firestore, 'staff'), where('id', '==', numId));
    const snap = await getDocs(q);
    snap.forEach(async (d) => {
      await updateDoc(d.ref, { deviceId });
    });
  } else {
    await db.staff.update(numId, { deviceId });
  }
}

export async function resetStaffDevice(staffId) {
  const numId = Number(staffId);
  if (isCloudActive) {
    const q = query(collection(firestore, 'staff'), where('id', '==', numId));
    const snap = await getDocs(q);
    snap.forEach(async (d) => {
      await updateDoc(d.ref, { deviceId: '' });
    });
  } else {
    await db.staff.update(numId, { deviceId: '' });
  }
}

// --- QUẢN LÝ NHÂN VIÊN ---
export async function getStaffList() {
  if (isCloudActive) {
    try {
      const snap = await getDocs(collection(firestore, 'staff'));
      return snap.docs.map(d => d.data());
    } catch (err) {
      console.error(err);
      return [];
    }
  } else {
    return await db.staff.toArray();
  }
}

export async function addStaff(name, email = '', username = '', password = '', role = 'employee', department = 'Ban Nội dung') {
  const newId = generateNumericId();
  const staffData = {
    id: newId,
    name,
    email,
    username,
    password,
    role,
    department,
    createdAt: new Date().toISOString()
  };

  if (isCloudActive) {
    await addDoc(collection(firestore, 'staff'), staffData);
    return newId;
  } else {
    return await db.staff.add(staffData);
  }
}

export async function updateStaff(id, name, email, username = '', password = '', role = 'employee', department = 'Ban Nội dung') {
  const numId = Number(id);
  if (isCloudActive) {
    const q = query(collection(firestore, 'staff'), where('id', '==', numId));
    const snap = await getDocs(q);
    snap.forEach(async (d) => {
      await updateDoc(d.ref, { name, email, username, password, role, department });
    });
  } else {
    await db.staff.update(numId, { name, email, username, password, role, department });
  }
}

export async function deleteStaff(id) {
  const numId = Number(id);
  if (isCloudActive) {
    const q = query(collection(firestore, 'staff'), where('id', '==', numId));
    const snap = await getDocs(q);
    snap.forEach(async (d) => {
      await deleteDoc(d.ref);
    });
  } else {
    await db.staff.delete(numId);
  }
}

// --- QUẢN LÝ KÊNH ---
export async function getChannelsList() {
  if (isCloudActive) {
    try {
      const snap = await getDocs(collection(firestore, 'channels'));
      return snap.docs.map(d => d.data());
    } catch (err) {
      console.error(err);
      return [];
    }
  } else {
    return await db.channels.toArray();
  }
}

export async function addChannel(name, link, assignedStaff = '', partner = '', startDate = '') {
  const newId = generateNumericId();
  const channelData = {
    id: newId,
    name,
    link,
    assignedStaff,
    partner,
    startDate,
    createdAt: new Date().toISOString()
  };

  if (isCloudActive) {
    await addDoc(collection(firestore, 'channels'), channelData);
    return newId;
  } else {
    return await db.channels.add(channelData);
  }
}

export async function updateChannel(id, name, link, assignedStaff, partner, startDate) {
  const numId = Number(id);
  if (isCloudActive) {
    const q = query(collection(firestore, 'channels'), where('id', '==', numId));
    const snap = await getDocs(q);
    snap.forEach(async (d) => {
      await updateDoc(d.ref, { name, link, assignedStaff, partner, startDate });
    });
  } else {
    await db.channels.update(numId, { name, link, assignedStaff, partner, startDate });
  }
}

export async function deleteChannel(id) {
  const numId = Number(id);
  if (isCloudActive) {
    // Xoá audits
    const qa = query(collection(firestore, 'audits'), where('channelId', '==', numId));
    const snapa = await getDocs(qa);
    snapa.forEach(async (d) => await deleteDoc(d.ref));

    // Xoá metrics
    const qm = query(collection(firestore, 'metrics'), where('channelId', '==', numId));
    const snapm = await getDocs(qm);
    snapm.forEach(async (d) => await deleteDoc(d.ref));

    // Xoá channel
    const qc = query(collection(firestore, 'channels'), where('id', '==', numId));
    const snapc = await getDocs(qc);
    snapc.forEach(async (d) => await deleteDoc(d.ref));
  } else {
    await db.audits.where('channelId').equals(numId).delete();
    await db.metrics.where('channelId').equals(numId).delete();
    await db.channels.delete(numId);
  }
}

// --- QUẢN LÝ SỐ LIỆU THÔ (METRICS) ---
export async function getMetricsList() {
  if (isCloudActive) {
    try {
      const snap = await getDocs(collection(firestore, 'metrics'));
      return snap.docs.map(d => d.data());
    } catch (err) {
      console.error(err);
      return [];
    }
  } else {
    return await db.metrics.toArray();
  }
}

export async function getMetricsByChannel(channelId) {
  const numId = Number(channelId);
  if (isCloudActive) {
    try {
      const q = query(collection(firestore, 'metrics'), where('channelId', '==', numId));
      const snap = await getDocs(q);
      return snap.docs.map(d => d.data());
    } catch (err) {
      console.error(err);
      return [];
    }
  } else {
    return await db.metrics.where('channelId').equals(numId).toArray();
  }
}

export async function addMetrics(metricsData) {
  const newId = generateNumericId();
  const payload = {
    ...metricsData,
    id: newId,
    createdAt: new Date().toISOString()
  };

  if (isCloudActive) {
    await addDoc(collection(firestore, 'metrics'), payload);
    return newId;
  } else {
    return await db.metrics.add(payload);
  }
}

export async function updateMetrics(id, metricsData) {
  const numId = Number(id);
  if (isCloudActive) {
    const q = query(collection(firestore, 'metrics'), where('id', '==', numId));
    const snap = await getDocs(q);
    snap.forEach(async (d) => {
      await updateDoc(d.ref, metricsData);
    });
  } else {
    await db.metrics.update(numId, metricsData);
  }
}

export async function deleteMetrics(id) {
  const numId = Number(id);
  if (isCloudActive) {
    const q = query(collection(firestore, 'metrics'), where('id', '==', numId));
    const snap = await getDocs(q);
    snap.forEach(async (d) => {
      await deleteDoc(d.ref);
    });
  } else {
    await db.metrics.delete(numId);
  }
}

// --- QUẢN LÝ BÁO CÁO ĐÁNH GIÁ (AUDITS) ---
export async function getAuditsList() {
  if (isCloudActive) {
    try {
      const snap = await getDocs(collection(firestore, 'audits'));
      return snap.docs.map(d => d.data());
    } catch (err) {
      console.error(err);
      return [];
    }
  } else {
    return await db.audits.toArray();
  }
}

export async function getAuditsByChannel(channelId) {
  const numId = Number(channelId);
  if (isCloudActive) {
    try {
      const q = query(collection(firestore, 'audits'), where('channelId', '==', numId));
      const snap = await getDocs(q);
      return snap.docs.map(d => d.data());
    } catch (err) {
      console.error(err);
      return [];
    }
  } else {
    return await db.audits.where('channelId').equals(numId).toArray();
  }
}

export async function addAudit(auditData) {
  const newId = generateNumericId();
  const payload = {
    ...auditData,
    id: newId,
    createdAt: new Date().toISOString()
  };

  if (isCloudActive) {
    await addDoc(collection(firestore, 'audits'), payload);
    return newId;
  } else {
    return await db.audits.add(payload);
  }
}

export async function updateAudit(id, auditData) {
  const numId = Number(id);
  if (isCloudActive) {
    const q = query(collection(firestore, 'audits'), where('id', '==', numId));
    const snap = await getDocs(q);
    snap.forEach(async (d) => {
      await updateDoc(d.ref, auditData);
    });
  } else {
    await db.audits.update(numId, auditData);
  }
}

export async function deleteAudit(id) {
  const numId = Number(id);
  if (isCloudActive) {
    const q = query(collection(firestore, 'audits'), where('id', '==', numId));
    const snap = await getDocs(q);
    snap.forEach(async (d) => {
      await deleteDoc(d.ref);
    });
  } else {
    await db.audits.delete(numId);
  }
}

// --- DỌN DẸP TOÀN BỘ CƠ SỞ DỮ LIỆU ---
export async function clearAllDatabaseData() {
  if (isCloudActive) {
    const collectionsToClear = ['channels', 'staff', 'metrics', 'audits'];
    for (const collName of collectionsToClear) {
      const snap = await getDocs(collection(firestore, collName));
      for (const d of snap.docs) {
        await deleteDoc(d.ref);
      }
    }
  } else {
    await db.channels.clear();
    await db.staff.clear();
    await db.metrics.clear();
    await db.audits.clear();
  }
}

// --- ĐỒNG BỘ DỮ LIỆU LOCAL LÊN FIREBASE CLOUD ---
export async function syncLocalToCloud() {
  if (!isCloudActive) {
    throw new Error('Chưa cấu hình Firebase Cloud!');
  }

  // Tải dữ liệu local
  const localStaff = await db.staff.toArray();
  const localChannels = await db.channels.toArray();
  const localMetrics = await db.metrics.toArray();
  const localAudits = await db.audits.toArray();

  // Đẩy Staff lên Cloud
  for (const item of localStaff) {
    const q = query(collection(firestore, 'staff'), where('id', '==', Number(item.id)));
    const snap = await getDocs(q);
    if (snap.empty) {
      await addDoc(collection(firestore, 'staff'), item);
    }
  }

  // Đẩy Channels lên Cloud
  for (const item of localChannels) {
    const q = query(collection(firestore, 'channels'), where('id', '==', Number(item.id)));
    const snap = await getDocs(q);
    if (snap.empty) {
      await addDoc(collection(firestore, 'channels'), item);
    }
  }

  // Đẩy Metrics lên Cloud
  for (const item of localMetrics) {
    const q = query(collection(firestore, 'metrics'), where('id', '==', Number(item.id)));
    const snap = await getDocs(q);
    if (snap.empty) {
      await addDoc(collection(firestore, 'metrics'), item);
    }
  }

  // Đẩy Audits lên Cloud
  for (const item of localAudits) {
    const q = query(collection(firestore, 'audits'), where('id', '==', Number(item.id)));
    const snap = await getDocs(q);
    if (snap.empty) {
      await addDoc(collection(firestore, 'audits'), item);
    }
  }
}

// --- SEED DỮ LIỆU ĐĂNG NHẬP MẪU CHUYÊN NGHIỆP ---
export async function seedDatabase() {
  const staffList = await getStaffList();
  const hasAdmin = staffList.some(s => s.username === 'admin');
  if (!hasAdmin) {
    console.log('🌱 Seeding default accounts...');
    // Thêm 3 tài khoản Demo với mật khẩu mặc định 100102
    await addStaff('Quản trị viên (Admin)', 'admin@ytaudit.com', 'admin', '100102', 'admin');
    await addStaff('Trưởng phòng (Manager)', 'truongphong@ytaudit.com', 'truongphong', '100102', 'manager');
    await addStaff('Nguyễn Văn A (Nhân viên)', 'nhanvien@ytaudit.com', 'nhanvien', '100102', 'employee');
    console.log('✅ Default accounts seeding completed!');
  }

  // Tự động dọn dẹp toàn bộ các kênh mẫu "Sample" cũ nếu tồn tại trong CSDL để trả lại giao diện sạch
  try {
    const channelsList = await getChannelsList();
    const sampleChannels = channelsList.filter(c => 
      c.name.includes('Sample') || 
      c.name.includes('Kênh Giải Trí Vui Vẻ') || 
      c.name.includes('Kênh Nhạc Thư Giãn')
    );
    for (const c of sampleChannels) {
      console.log(`🗑️ Tự động xóa kênh mẫu: ${c.name}`);
      await deleteChannel(c.id);
    }
  } catch (err) {
    console.error('Lỗi khi tự động dọn dẹp kênh mẫu:', err);
  }
}
