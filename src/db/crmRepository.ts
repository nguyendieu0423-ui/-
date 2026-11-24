import { and, asc, eq, or, sql } from 'drizzle-orm';
import { db, markDatabaseUnavailable } from './index.ts';

function isConnectivityErr(error: unknown): boolean {
  const err = error as { code?: string; message?: string; cause?: { code?: string; message?: string } };
  const code = err?.code || err?.cause?.code || '';
  const message = `${err?.message || ''} ${err?.cause?.message || ''}`;
  return (
    code === 'ENOENT' ||
    code === 'ECONNREFUSED' ||
    code === 'ENOTFOUND' ||
    code === 'ETIMEDOUT' ||
    message.includes('ENOENT') ||
    message.includes('Connection terminated unexpectedly')
  );
}
import {
  candidates,
  chatMessages,
  kyujins,
  ptPartners,
  resignations,
} from './schema.ts';
import type {
  Candidate,
  KyujinCompany,
  PtPartner,
  ResignationRecord,
} from '../types/crm.ts';

export interface MessengerExtractedProfile {
  name?: string;
  jlpt?: string;
  phone?: string;
  desiredIndustry?: string;
  desiredArea?: string;
  visaExpiry?: string;
}

export interface CrmSnapshot {
  candidates: Candidate[];
  schoolStudents: Candidate[];
  kyujins: KyujinCompany[];
  ptPartners: PtPartner[];
  resignations?: ResignationRecord[];
}

export async function loadUserCrmData(userId: number): Promise<
  CrmSnapshot & { hasExistingData: boolean }
> {
  try {
    const [candRows, kyujinRows, ptRows, resigRows] = await Promise.all([
      db
        .select()
        .from(candidates)
        .where(eq(candidates.userId, userId))
        .orderBy(asc(candidates.id)),
      db
        .select()
        .from(kyujins)
        .where(eq(kyujins.userId, userId))
        .orderBy(asc(kyujins.id)),
      db
        .select()
        .from(ptPartners)
        .where(eq(ptPartners.userId, userId))
        .orderBy(asc(ptPartners.id)),
      db
        .select()
        .from(resignations)
        .where(eq(resignations.userId, userId))
        .orderBy(asc(resignations.id)),
    ]);

    const regularCandidates: Candidate[] = [];
    const schoolStudents: Candidate[] = [];

    for (const row of candRows) {
      if (row.isSchool) {
        schoolStudents.push(row.payload);
      } else {
        regularCandidates.push(row.payload);
      }
    }

    return {
      candidates: regularCandidates,
      schoolStudents,
      kyujins: kyujinRows.map((r) => r.payload),
      ptPartners: ptRows.map((r) => r.payload),
      resignations: resigRows.map((r) => r.payload),
      hasExistingData:
        candRows.length > 0 ||
        kyujinRows.length > 0 ||
        ptRows.length > 0 ||
        resigRows.length > 0,
    };
  } catch (error) {
    if (isConnectivityErr(error)) {
      markDatabaseUnavailable();
    } else {
      console.error('Database query failed in loadUserCrmData:', error);
    }
    throw new Error('Failed to load CRM records from database.', {
      cause: error,
    });
  }
}

export async function saveUserCrmSnapshot(
  userId: number,
  snapshot: CrmSnapshot
): Promise<void> {
  try {
    await db.transaction(async (tx) => {
      await tx.delete(candidates).where(eq(candidates.userId, userId));
      await tx.delete(kyujins).where(eq(kyujins.userId, userId));
      await tx.delete(ptPartners).where(eq(ptPartners.userId, userId));
      if (Array.isArray(snapshot.resignations)) {
        await tx.delete(resignations).where(eq(resignations.userId, userId));
      }

      const combinedCandidates = [
        ...snapshot.candidates.map((c) => ({
          userId,
          recordId: `cand_${String(c.id)}`,
          name: String(c.name || 'Unnamed'),
          fbName: c.fbName || null,
          channel: String(c.channel || 'company'),
          isSchool: false,
          status: String(c.progressStatus || '新規相談'),
          jlpt: c.jlpt || null,
          visaStatus: c.visaStatus || null,
          visaExpiry: c.visaExpiry || null,
          industry: c.desiredIndustry || null,
          area: c.desiredArea || null,
          conversionGoal: c.conversionGoal || null,
          negotiationLog: c.negotiationLog || null,
          proposedCompany: c.proposedCompany || null,
          googleCalendarEventId: c.googleCalendarEventId || null,
          payload: c,
        })),
        ...snapshot.schoolStudents.map((s) => ({
          userId,
          recordId: `school_${String(s.id)}`,
          name: String(s.name || 'Unnamed'),
          fbName: s.fbName || null,
          channel: 'school',
          isSchool: true,
          status: String(s.progressStatus || '新規相談'),
          jlpt: s.jlpt || null,
          visaStatus: s.visaStatus || null,
          visaExpiry: s.visaExpiry || null,
          industry: s.desiredIndustry || null,
          area: s.desiredArea || null,
          conversionGoal: s.conversionGoal || null,
          negotiationLog: s.negotiationLog || null,
          proposedCompany: s.proposedCompany || null,
          googleCalendarEventId: s.googleCalendarEventId || null,
          payload: s,
        })),
      ];

      if (combinedCandidates.length > 0) {
        await tx.insert(candidates).values(combinedCandidates);
      }

      if (snapshot.kyujins.length > 0) {
        await tx.insert(kyujins).values(
          snapshot.kyujins.map((k) => ({
            userId,
            recordId: String(k.id),
            companyName: String(k.companyName || 'Unnamed Company'),
            sector: k.sector || k.areaBreakdown?.[0]?.sector || null,
            industry: k.industry || null,
            employmentType:
              k.employmentType || k.areaBreakdown?.[0]?.employmentType || null,
            area: k.areaBreakdown?.[0]?.area || null,
            payload: k,
          }))
        );
      }

      if (snapshot.ptPartners.length > 0) {
        await tx.insert(ptPartners).values(
          snapshot.ptPartners.map((p) => ({
            userId,
            recordId: String(p.id),
            name: String(p.partnerName || 'Unnamed Partner'),
            payload: p,
          }))
        );
      }

      if (Array.isArray(snapshot.resignations) && snapshot.resignations.length > 0) {
        await tx.insert(resignations).values(
          snapshot.resignations.map((r) => ({
            userId,
            recordId: String(r.id),
            fullName: String(r.fullName || 'Unnamed'),
            nationality: r.nationality || null,
            previousCompany: String(r.previousCompany || 'Chưa rõ công ty'),
            resignationDate: r.resignationDate || null,
            paperworkDeadline: r.paperworkDeadline || null,
            paperworkStatus: String(r.paperworkStatus || 'Chưa làm'),
            handoverNotes: r.handoverNotes || null,
            newJobNeed: r.newJobNeed || null,
            googleCalendarEventId: r.googleCalendarEventId || null,
            payload: r,
          }))
        );
      }
    });
  } catch (error) {
    if (isConnectivityErr(error)) {
      markDatabaseUnavailable();
    } else {
      console.error('Database write failed in saveUserCrmSnapshot:', error);
    }
    throw new Error('Failed to save CRM records to database.', {
      cause: error,
    });
  }
}

export async function getUserChatHistory(
  userId: number,
  sessionId = 'default'
) {
  try {
    return await db
      .select({
        id: chatMessages.id,
        role: chatMessages.role,
        assistantRole: chatMessages.assistantRole,
        modelUsed: chatMessages.modelUsed,
        content: chatMessages.content,
        createdAt: chatMessages.createdAt,
      })
      .from(chatMessages)
      .where(
        and(
          eq(chatMessages.userId, userId),
          eq(chatMessages.sessionId, sessionId)
        )
      )
      .orderBy(asc(chatMessages.id));
  } catch (error) {
    if (isConnectivityErr(error)) {
      markDatabaseUnavailable();
    } else {
      console.error('Database query failed in getUserChatHistory:', error);
    }
    throw new Error('Failed to load conversation history.', { cause: error });
  }
}

export async function appendUserChatExchange(
  userId: number,
  sessionId: string,
  assistantRole: string,
  modelUsed: string,
  userContent: string,
  modelContent: string
) {
  try {
    return await db
      .insert(chatMessages)
      .values([
        {
          userId,
          sessionId,
          role: 'user',
          assistantRole,
          modelUsed,
          content: userContent,
        },
        {
          userId,
          sessionId,
          role: 'model',
          assistantRole,
          modelUsed,
          content: modelContent,
        },
      ])
      .returning();
  } catch (error) {
    if (isConnectivityErr(error)) {
      markDatabaseUnavailable();
    } else {
      console.error('Database insert failed in appendUserChatExchange:', error);
    }
    throw new Error('Failed to save conversation turn.', { cause: error });
  }
}

export async function clearUserChatHistory(
  userId: number,
  sessionId = 'default'
) {
  try {
    await db
      .delete(chatMessages)
      .where(
        and(
          eq(chatMessages.userId, userId),
          eq(chatMessages.sessionId, sessionId)
        )
      );
  } catch (error) {
    if (isConnectivityErr(error)) {
      markDatabaseUnavailable();
    } else {
      console.error('Database delete failed in clearUserChatHistory:', error);
    }
    throw new Error('Failed to clear conversation history.', { cause: error });
  }
}

function cleanExtractedField(value?: string): string {
  const trimmed = String(value || '').trim();
  if (
    !trimmed ||
    trimmed === '不明' ||
    trimmed === 'なし' ||
    trimmed === 'null' ||
    trimmed === 'undefined' ||
    trimmed === '-'
  ) {
    return '';
  }
  return trimmed;
}

function isPlaceholderValue(value?: string): boolean {
  const v = String(value || '').trim();
  return (
    !v ||
    v === '未確認' ||
    v === '未定' ||
    v === '不明' ||
    v === 'Unnamed' ||
    v.startsWith('FB相談者 (')
  );
}

export async function upsertMessengerCandidate(
  userId: number,
  input: {
    senderId: string;
    rawText: string;
    extracted: MessengerExtractedProfile;
  }
): Promise<{ action: 'created' | 'updated'; candidate: Candidate }> {
  const senderId = String(input.senderId || '').trim();
  const rawText = String(input.rawText || '').trim();
  if (!senderId || !rawText) {
    throw new Error('Missing senderId or message text.');
  }

  const now = new Date();
  const jstDate = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const yyyy = jstDate.getUTCFullYear();
  const mm = String(jstDate.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(jstDate.getUTCDate()).padStart(2, '0');
  const hh = String(jstDate.getUTCHours()).padStart(2, '0');
  const min = String(jstDate.getUTCMinutes()).padStart(2, '0');
  const todayYmd = `${yyyy}-${mm}-${dd}`;
  const timestamp = `${todayYmd} ${hh}:${min}`;
  const memoEntry = `[Messenger ${timestamp}] ${rawText}`;

  const extName = cleanExtractedField(input.extracted.name);
  const extJlpt = cleanExtractedField(input.extracted.jlpt);
  const extPhone = cleanExtractedField(input.extracted.phone);
  const extIndustry = cleanExtractedField(input.extracted.desiredIndustry);
  const extArea = cleanExtractedField(input.extracted.desiredArea);
  const extVisaExpiry = cleanExtractedField(input.extracted.visaExpiry);

  const canonicalId = `fb_${senderId}`;
  const canonicalRecordId = `cand_${canonicalId}`;

  try {
    return await db.transaction(async (tx) => {
      const existingRows = await tx
        .select()
        .from(candidates)
        .where(
          and(
            eq(candidates.userId, userId),
            or(
              eq(candidates.recordId, canonicalRecordId),
              sql`${candidates.payload}->>'fbSenderId' = ${senderId}`
            )
          )
        )
        .limit(1);

      if (existingRows[0]) {
        const existingRow = existingRows[0];
        const prevPayload = existingRow.payload;
        const prevMemo = String(prevPayload.memo || '').trim();
        const updatedMemo = prevMemo
          ? `${prevMemo}\n${memoEntry}`
          : memoEntry;

        const nextName =
          isPlaceholderValue(prevPayload.name) && extName
            ? extName
            : prevPayload.name || `FB相談者 (${senderId.slice(-6)})`;
        const nextJlpt =
          isPlaceholderValue(prevPayload.jlpt) && extJlpt
            ? extJlpt
            : prevPayload.jlpt || '';
        const nextPhone =
          isPlaceholderValue(prevPayload.phone) && extPhone
            ? extPhone
            : prevPayload.phone || '';
        const nextIndustry =
          isPlaceholderValue(prevPayload.desiredIndustry) && extIndustry
            ? extIndustry
            : prevPayload.desiredIndustry || '';
        const nextArea =
          isPlaceholderValue(prevPayload.desiredArea) && extArea
            ? extArea
            : prevPayload.desiredArea || '';
        const nextVisaExpiry =
          isPlaceholderValue(prevPayload.visaExpiry) && extVisaExpiry
            ? extVisaExpiry
            : prevPayload.visaExpiry || '';

        const nextPayload: Candidate = {
          ...prevPayload,
          id: prevPayload.id || canonicalId,
          fbSenderId: senderId,
          name: nextName,
          fbName:
            prevPayload.fbName ||
            extName ||
            `Messenger ${senderId.slice(-6)}`,
          channel: prevPayload.channel || 'company',
          jlpt: nextJlpt,
          phone: nextPhone,
          desiredIndustry: nextIndustry,
          desiredArea: nextArea,
          visaExpiry: nextVisaExpiry,
          lastContactDate: todayYmd,
          memo: updatedMemo,
        };

        await tx
          .update(candidates)
          .set({
            name: String(nextPayload.name || 'Unnamed'),
            jlpt: nextPayload.jlpt || null,
            industry: nextPayload.desiredIndustry || null,
            area: nextPayload.desiredArea || null,
            payload: nextPayload,
            updatedAt: new Date(),
          })
          .where(eq(candidates.id, existingRow.id));

        return { action: 'updated', candidate: nextPayload };
      }

      const newCandidate: Candidate = {
        id: canonicalId,
        fbSenderId: senderId,
        name: extName || `FB相談者 (${senderId.slice(-6)})`,
        fbName: extName || `Messenger ${senderId.slice(-6)}`,
        channel: 'company',
        progressStatus: '新規相談',
        jlpt: extJlpt,
        phone: extPhone,
        desiredIndustry: extIndustry,
        desiredArea: extArea,
        visaExpiry: extVisaExpiry,
        lastContactDate: todayYmd,
        memo: memoEntry,
      };

      await tx.insert(candidates).values({
        userId,
        recordId: canonicalRecordId,
        name: String(newCandidate.name),
        channel: 'company',
        isSchool: false,
        status: '新規相談',
        jlpt: newCandidate.jlpt || null,
        industry: newCandidate.desiredIndustry || null,
        area: newCandidate.desiredArea || null,
        proposedCompany: null,
        payload: newCandidate,
        updatedAt: new Date(),
      });

      return { action: 'created', candidate: newCandidate };
    });
  } catch (error) {
    if (isConnectivityErr(error)) {
      markDatabaseUnavailable();
    } else {
      console.error('Database upsert failed in upsertMessengerCandidate:', error);
    }
    throw new Error('Failed to store Messenger candidate record.', {
      cause: error,
    });
  }
}

export async function updateCandidateCalendarEvent(
  userId: number,
  candidateId: string,
  eventId: string | null,
  eventLink: string | null
): Promise<void> {
  const safeId = String(candidateId || '').trim();
  if (!safeId) return;

  try {
    const matchingRows = await db
      .select()
      .from(candidates)
      .where(
        and(
          eq(candidates.userId, userId),
          or(
            eq(candidates.recordId, `cand_${safeId}`),
            eq(candidates.recordId, `school_${safeId}`)
          )
        )
      )
      .limit(1);

    const row = matchingRows[0];
    if (!row) return;

    const updatedPayload: Candidate = {
      ...row.payload,
      googleCalendarEventId: eventId || undefined,
      googleCalendarEventLink: eventLink || undefined,
    };

    await db
      .update(candidates)
      .set({
        googleCalendarEventId: eventId,
        payload: updatedPayload,
        updatedAt: new Date(),
      })
      .where(eq(candidates.id, row.id));
  } catch (error) {
    if (isConnectivityErr(error)) {
      markDatabaseUnavailable();
    } else {
      console.error(
        'Database update failed in updateCandidateCalendarEvent:',
        error
      );
    }
    throw new Error('Failed to update Google Calendar event link in DB.', {
      cause: error,
    });
  }
}

export async function updateResignationCalendarEvent(
  userId: number,
  recordId: string,
  eventId: string | null,
  eventLink: string | null
): Promise<void> {
  const safeId = String(recordId || '').trim();
  if (!safeId) return;

  try {
    const matchingRows = await db
      .select()
      .from(resignations)
      .where(
        and(
          eq(resignations.userId, userId),
          eq(resignations.recordId, safeId)
        )
      )
      .limit(1);

    const row = matchingRows[0];
    if (!row) return;

    const updatedPayload: ResignationRecord = {
      ...row.payload,
      googleCalendarEventId: eventId || undefined,
      googleCalendarEventLink: eventLink || undefined,
    };

    await db
      .update(resignations)
      .set({
        googleCalendarEventId: eventId,
        payload: updatedPayload,
        updatedAt: new Date(),
      })
      .where(eq(resignations.id, row.id));
  } catch (error) {
    if (isConnectivityErr(error)) {
      markDatabaseUnavailable();
    } else {
      console.error(
        'Database update failed in updateResignationCalendarEvent:',
        error
      );
    }
    throw new Error('Failed to update resignation Google Calendar event link in DB.', {
      cause: error,
    });
  }
}



