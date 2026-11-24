import { relations } from 'drizzle-orm';
import {
  boolean,
  integer,
  jsonb,
  pgTable,
  serial,
  text,
  timestamp,
  uniqueIndex,
} from 'drizzle-orm/pg-core';
import type {
  Candidate,
  KyujinCompany,
  PtPartner,
  ResignationRecord,
} from '../types/crm.ts';

export const users = pgTable('users', {
  id: serial('id').primaryKey(),
  uid: text('uid').notNull().unique(),
  email: text('email').notNull(),
  displayName: text('display_name'),
  createdAt: timestamp('created_at').defaultNow().notNull(),
});

export const candidates = pgTable(
  'candidates',
  {
    id: serial('id').primaryKey(),
    userId: integer('user_id')
      .references(() => users.id, { onDelete: 'cascade' })
      .notNull(),
    recordId: text('record_id').notNull(),
    name: text('name').notNull(),
    fbName: text('fb_name'),
    channel: text('channel').notNull(),
    isSchool: boolean('is_school').default(false).notNull(),
    status: text('status').default('未対応').notNull(),
    jlpt: text('jlpt'),
    visaStatus: text('visa_status'),
    visaExpiry: text('visa_expiry'),
    industry: text('industry'),
    area: text('area'),
    conversionGoal: text('conversion_goal'),
    negotiationLog: text('negotiation_log'),
    proposedCompany: text('proposed_company'),
    googleCalendarEventId: text('google_calendar_event_id'),
    payload: jsonb('payload').$type<Candidate>().notNull(),
    updatedAt: timestamp('updated_at').defaultNow().notNull(),
  },
  (table) => [
    uniqueIndex('candidates_user_record_idx').on(table.userId, table.recordId),
  ]
);

export const kyujins = pgTable(
  'kyujins',
  {
    id: serial('id').primaryKey(),
    userId: integer('user_id')
      .references(() => users.id, { onDelete: 'cascade' })
      .notNull(),
    recordId: text('record_id').notNull(),
    companyName: text('company_name').notNull(),
    sector: text('sector'),
    industry: text('industry'),
    employmentType: text('employment_type'),
    area: text('area'),
    payload: jsonb('payload').$type<KyujinCompany>().notNull(),
    updatedAt: timestamp('updated_at').defaultNow().notNull(),
  },
  (table) => [
    uniqueIndex('kyujins_user_record_idx').on(table.userId, table.recordId),
  ]
);

export const ptPartners = pgTable(
  'pt_partners',
  {
    id: serial('id').primaryKey(),
    userId: integer('user_id')
      .references(() => users.id, { onDelete: 'cascade' })
      .notNull(),
    recordId: text('record_id').notNull(),
    name: text('name').notNull(),
    payload: jsonb('payload').$type<PtPartner>().notNull(),
    updatedAt: timestamp('updated_at').defaultNow().notNull(),
  },
  (table) => [
    uniqueIndex('pt_partners_user_record_idx').on(table.userId, table.recordId),
  ]
);

export const resignations = pgTable(
  'resignations',
  {
    id: serial('id').primaryKey(),
    userId: integer('user_id')
      .references(() => users.id, { onDelete: 'cascade' })
      .notNull(),
    recordId: text('record_id').notNull(),
    fullName: text('full_name').notNull(),
    nationality: text('nationality'),
    previousCompany: text('previous_company').notNull(),
    resignationDate: text('resignation_date'),
    paperworkDeadline: text('paperwork_deadline'),
    paperworkStatus: text('paperwork_status').default('Chưa làm').notNull(),
    handoverNotes: text('handover_notes'),
    newJobNeed: text('new_job_need'),
    googleCalendarEventId: text('google_calendar_event_id'),
    payload: jsonb('payload').$type<ResignationRecord>().notNull(),
    updatedAt: timestamp('updated_at').defaultNow().notNull(),
  },
  (table) => [
    uniqueIndex('resignations_user_record_idx').on(table.userId, table.recordId),
  ]
);

export const chatMessages = pgTable('chat_messages', {
  id: serial('id').primaryKey(),
  userId: integer('user_id')
    .references(() => users.id, { onDelete: 'cascade' })
    .notNull(),
  sessionId: text('session_id').default('default').notNull(),
  role: text('role').notNull(),
  assistantRole: text('assistant_role').default('crm_strategist').notNull(),
  modelUsed: text('model_used').default('gemini-4-hydrogen').notNull(),
  content: text('content').notNull(),
  createdAt: timestamp('created_at').defaultNow().notNull(),
});

export const usersRelations = relations(users, ({ many }) => ({
  candidates: many(candidates),
  kyujins: many(kyujins),
  ptPartners: many(ptPartners),
  resignations: many(resignations),
  chatMessages: many(chatMessages),
}));

export const candidatesRelations = relations(candidates, ({ one }) => ({
  owner: one(users, {
    fields: [candidates.userId],
    references: [users.id],
  }),
}));

export const kyujinsRelations = relations(kyujins, ({ one }) => ({
  owner: one(users, {
    fields: [kyujins.userId],
    references: [users.id],
  }),
}));

export const ptPartnersRelations = relations(ptPartners, ({ one }) => ({
  owner: one(users, {
    fields: [ptPartners.userId],
    references: [users.id],
  }),
}));

export const resignationsRelations = relations(resignations, ({ one }) => ({
  owner: one(users, {
    fields: [resignations.userId],
    references: [users.id],
  }),
}));

export const chatMessagesRelations = relations(chatMessages, ({ one }) => ({
  owner: one(users, {
    fields: [chatMessages.userId],
    references: [users.id],
  }),
}));

