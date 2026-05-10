--
-- PostgreSQL database dump
--


-- Dumped from database version 17.7 (Debian 17.7-3.pgdg13+1)
-- Dumped by pg_dump version 18.3

SET statement_timeout = 0;
SET lock_timeout = 0;
SET idle_in_transaction_session_timeout = 0;
SET transaction_timeout = 0;
SET client_encoding = 'UTF8';
SET standard_conforming_strings = on;
SELECT pg_catalog.set_config('search_path', '', false);
SET check_function_bodies = false;
SET xmloption = content;
SET client_min_messages = warning;
SET row_security = off;


BEGIN;

ALTER TABLE IF EXISTS ONLY "public"."_SubjectToUser" DROP CONSTRAINT IF EXISTS "_SubjectToUser_B_fkey";
ALTER TABLE IF EXISTS ONLY "public"."_SubjectToUser" DROP CONSTRAINT IF EXISTS "_SubjectToUser_A_fkey";
ALTER TABLE IF EXISTS ONLY "public"."_RoleToUser" DROP CONSTRAINT IF EXISTS "_RoleToUser_B_fkey";
ALTER TABLE IF EXISTS ONLY "public"."_RoleToUser" DROP CONSTRAINT IF EXISTS "_RoleToUser_A_fkey";
ALTER TABLE IF EXISTS ONLY "public"."_ClassToUser" DROP CONSTRAINT IF EXISTS "_ClassToUser_B_fkey";
ALTER TABLE IF EXISTS ONLY "public"."_ClassToUser" DROP CONSTRAINT IF EXISTS "_ClassToUser_A_fkey";
ALTER TABLE IF EXISTS ONLY "public"."PupilCode" DROP CONSTRAINT IF EXISTS "PupilCode_groupId_fkey";
ALTER TABLE IF EXISTS ONLY "public"."PupilCode" DROP CONSTRAINT IF EXISTS "PupilCode_assignmentId_fkey";
ALTER TABLE IF EXISTS ONLY "public"."IgnoredWord" DROP CONSTRAINT IF EXISTS "IgnoredWord_teacherId_fkey";
ALTER TABLE IF EXISTS ONLY "public"."CommonPupilCode" DROP CONSTRAINT IF EXISTS "CommonPupilCode_commonGroupId_fkey";
ALTER TABLE IF EXISTS ONLY "public"."CommonPupilCode" DROP CONSTRAINT IF EXISTS "CommonPupilCode_assignmentId_fkey";
ALTER TABLE IF EXISTS ONLY "public"."CommonCommentOption" DROP CONSTRAINT IF EXISTS "CommonCommentOption_groupId_fkey";
ALTER TABLE IF EXISTS ONLY "public"."CommentOption" DROP CONSTRAINT IF EXISTS "CommentOption_groupId_fkey";
ALTER TABLE IF EXISTS ONLY "public"."CommentGroup" DROP CONSTRAINT IF EXISTS "CommentGroup_subjectId_fkey";
ALTER TABLE IF EXISTS ONLY "public"."Class" DROP CONSTRAINT IF EXISTS "Class_subjectId_fkey";
ALTER TABLE IF EXISTS ONLY "public"."Assignment" DROP CONSTRAINT IF EXISTS "Assignment_pupilId_fkey";
ALTER TABLE IF EXISTS ONLY "public"."Assignment" DROP CONSTRAINT IF EXISTS "Assignment_classId_fkey";
ALTER TABLE IF EXISTS ONLY "public"."Assignment" DROP CONSTRAINT IF EXISTS "Assignment_checkedById_fkey";
DROP INDEX IF EXISTS "public"."_SubjectToUser_B_index";
DROP INDEX IF EXISTS "public"."_RoleToUser_B_index";
DROP INDEX IF EXISTS "public"."_ClassToUser_B_index";
DROP INDEX IF EXISTS "public"."User_username_key";
DROP INDEX IF EXISTS "public"."Subject_code_key";
DROP INDEX IF EXISTS "public"."Role_name_key";
DROP INDEX IF EXISTS "public"."PupilCode_groupId_idx";
DROP INDEX IF EXISTS "public"."PupilCode_assignmentId_idx";
DROP INDEX IF EXISTS "public"."PupilCode_assignmentId_groupId_key";
DROP INDEX IF EXISTS "public"."IgnoredWord_teacherId_idx";
DROP INDEX IF EXISTS "public"."CommonPupilCode_commonGroupId_idx";
DROP INDEX IF EXISTS "public"."CommonPupilCode_assignmentId_idx";
DROP INDEX IF EXISTS "public"."CommonPupilCode_assignmentId_commonGroupId_key";
DROP INDEX IF EXISTS "public"."CommonCommentOption_groupId_idx";
DROP INDEX IF EXISTS "public"."CommonCommentOption_groupId_code_key";
DROP INDEX IF EXISTS "public"."CommonCommentGroup_name_key";
DROP INDEX IF EXISTS "public"."CommentOption_groupId_idx";
DROP INDEX IF EXISTS "public"."CommentOption_groupId_code_key";
DROP INDEX IF EXISTS "public"."CommentGroup_subjectId_name_key";
DROP INDEX IF EXISTS "public"."Class_subjectId_idx";
DROP INDEX IF EXISTS "public"."Class_name_key";
DROP INDEX IF EXISTS "public"."AuditLog_userId_idx";
DROP INDEX IF EXISTS "public"."AuditLog_entityType_idx";
DROP INDEX IF EXISTS "public"."AuditLog_createdAt_idx";
DROP INDEX IF EXISTS "public"."AuditLog_action_idx";
DROP INDEX IF EXISTS "public"."Assignment_pupilId_idx";
DROP INDEX IF EXISTS "public"."Assignment_classId_idx";
DROP INDEX IF EXISTS "public"."Assignment_checkStatus_idx";
ALTER TABLE IF EXISTS ONLY "public"."_prisma_migrations" DROP CONSTRAINT IF EXISTS "_prisma_migrations_pkey";
ALTER TABLE IF EXISTS ONLY "public"."_SubjectToUser" DROP CONSTRAINT IF EXISTS "_SubjectToUser_AB_pkey";
ALTER TABLE IF EXISTS ONLY "public"."_RoleToUser" DROP CONSTRAINT IF EXISTS "_RoleToUser_AB_pkey";
ALTER TABLE IF EXISTS ONLY "public"."_ClassToUser" DROP CONSTRAINT IF EXISTS "_ClassToUser_AB_pkey";
ALTER TABLE IF EXISTS ONLY "public"."User" DROP CONSTRAINT IF EXISTS "User_pkey";
ALTER TABLE IF EXISTS ONLY "public"."Subject" DROP CONSTRAINT IF EXISTS "Subject_pkey";
ALTER TABLE IF EXISTS ONLY "public"."Role" DROP CONSTRAINT IF EXISTS "Role_pkey";
ALTER TABLE IF EXISTS ONLY "public"."Pupil" DROP CONSTRAINT IF EXISTS "Pupil_pkey";
ALTER TABLE IF EXISTS ONLY "public"."PupilCode" DROP CONSTRAINT IF EXISTS "PupilCode_pkey";
ALTER TABLE IF EXISTS ONLY "public"."IgnoredWord" DROP CONSTRAINT IF EXISTS "IgnoredWord_teacherId_word_key";
ALTER TABLE IF EXISTS ONLY "public"."IgnoredWord" DROP CONSTRAINT IF EXISTS "IgnoredWord_pkey";
ALTER TABLE IF EXISTS ONLY "public"."Deadline" DROP CONSTRAINT IF EXISTS "Deadline_pkey";
ALTER TABLE IF EXISTS ONLY "public"."CommonPupilCode" DROP CONSTRAINT IF EXISTS "CommonPupilCode_pkey";
ALTER TABLE IF EXISTS ONLY "public"."CommonCommentOption" DROP CONSTRAINT IF EXISTS "CommonCommentOption_pkey";
ALTER TABLE IF EXISTS ONLY "public"."CommonCommentGroup" DROP CONSTRAINT IF EXISTS "CommonCommentGroup_pkey";
ALTER TABLE IF EXISTS ONLY "public"."CommentOption" DROP CONSTRAINT IF EXISTS "CommentOption_pkey";
ALTER TABLE IF EXISTS ONLY "public"."CommentGroup" DROP CONSTRAINT IF EXISTS "CommentGroup_pkey";
ALTER TABLE IF EXISTS ONLY "public"."Class" DROP CONSTRAINT IF EXISTS "Class_pkey";
ALTER TABLE IF EXISTS ONLY "public"."AuditLog" DROP CONSTRAINT IF EXISTS "AuditLog_pkey";
ALTER TABLE IF EXISTS ONLY "public"."Assignment" DROP CONSTRAINT IF EXISTS "Assignment_pkey";
ALTER TABLE IF EXISTS ONLY "public"."AppSetting" DROP CONSTRAINT IF EXISTS "AppSetting_pkey";
DROP TABLE IF EXISTS "public"."_prisma_migrations";
DROP TABLE IF EXISTS "public"."_SubjectToUser";
DROP TABLE IF EXISTS "public"."_RoleToUser";
DROP TABLE IF EXISTS "public"."_ClassToUser";
DROP TABLE IF EXISTS "public"."User";
DROP TABLE IF EXISTS "public"."Subject";
DROP TABLE IF EXISTS "public"."Role";
DROP TABLE IF EXISTS "public"."PupilCode";
DROP TABLE IF EXISTS "public"."Pupil";
DROP TABLE IF EXISTS "public"."IgnoredWord";
DROP TABLE IF EXISTS "public"."Deadline";
DROP TABLE IF EXISTS "public"."CommonPupilCode";
DROP TABLE IF EXISTS "public"."CommonCommentOption";
DROP TABLE IF EXISTS "public"."CommonCommentGroup";
DROP TABLE IF EXISTS "public"."CommentOption";
DROP TABLE IF EXISTS "public"."CommentGroup";
DROP TABLE IF EXISTS "public"."Class";
DROP TABLE IF EXISTS "public"."AuditLog";
DROP TABLE IF EXISTS "public"."Assignment";
DROP TABLE IF EXISTS "public"."AppSetting";
--
-- Name: SCHEMA "public"; Type: COMMENT; Schema: -; Owner: -
--

COMMENT ON SCHEMA "public" IS 'standard public schema';


SET default_tablespace = '';

SET default_table_access_method = "heap";

--
-- Name: AppSetting; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."AppSetting" (
    "key" "text" NOT NULL,
    "value" "text" NOT NULL
);


--
-- Name: Assignment; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."Assignment" (
    "id" "text" NOT NULL,
    "pupilId" "text" NOT NULL,
    "classId" "text" NOT NULL,
    "eoyLevel" "text",
    "targetLevel" "text",
    "actualLevel" "text",
    "finalComment" "text",
    "checkNote" "text",
    "checkStatus" "text" DEFAULT 'not_required'::"text" NOT NULL,
    "checkedAt" timestamp(3) without time zone,
    "checkedById" "text",
    "linkedData" "jsonb",
    "aiStage" "text"
);


--
-- Name: AuditLog; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."AuditLog" (
    "id" "text" NOT NULL,
    "userId" "text",
    "username" "text",
    "action" "text" NOT NULL,
    "entityType" "text",
    "entityId" "text",
    "details" "text",
    "ipAddress" "text",
    "userAgent" "text",
    "createdAt" timestamp(3) without time zone DEFAULT CURRENT_TIMESTAMP NOT NULL
);


--
-- Name: Class; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."Class" (
    "id" "text" NOT NULL,
    "name" "text" NOT NULL,
    "year" "text",
    "subjectId" "text" NOT NULL
);


--
-- Name: CommentGroup; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."CommentGroup" (
    "id" "text" NOT NULL,
    "name" "text" NOT NULL,
    "displayOrder" integer DEFAULT 0 NOT NULL,
    "subjectId" "text" NOT NULL,
    "title" "text" DEFAULT ''::"text" NOT NULL,
    "isLinked" boolean DEFAULT false NOT NULL,
    "linkedField" "text"
);


--
-- Name: CommentOption; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."CommentOption" (
    "id" "text" NOT NULL,
    "code" "text" NOT NULL,
    "text" "text" NOT NULL,
    "displayOrder" integer DEFAULT 0 NOT NULL,
    "groupId" "text" NOT NULL
);


--
-- Name: CommonCommentGroup; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."CommonCommentGroup" (
    "id" "text" NOT NULL,
    "name" "text" NOT NULL,
    "title" "text" DEFAULT ''::"text" NOT NULL,
    "displayOrder" integer DEFAULT 0 NOT NULL,
    "isLinked" boolean DEFAULT false NOT NULL,
    "linkedField" "text"
);


--
-- Name: CommonCommentOption; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."CommonCommentOption" (
    "id" "text" NOT NULL,
    "code" "text" NOT NULL,
    "text" "text" NOT NULL,
    "displayOrder" integer DEFAULT 0 NOT NULL,
    "groupId" "text" NOT NULL
);


--
-- Name: CommonPupilCode; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."CommonPupilCode" (
    "id" "text" NOT NULL,
    "assignmentId" "text" NOT NULL,
    "commonGroupId" "text" NOT NULL,
    "code" "text"
);


--
-- Name: Deadline; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."Deadline" (
    "id" "text" NOT NULL,
    "title" "text" NOT NULL,
    "date" timestamp(3) without time zone NOT NULL,
    "description" "text",
    "isActive" boolean DEFAULT true NOT NULL,
    "createdAt" timestamp(3) without time zone DEFAULT CURRENT_TIMESTAMP NOT NULL
);


--
-- Name: IgnoredWord; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."IgnoredWord" (
    "id" "text" DEFAULT ("gen_random_uuid"())::"text" NOT NULL,
    "teacherId" "text" NOT NULL,
    "word" "text" NOT NULL,
    "createdAt" timestamp with time zone DEFAULT "now"()
);


--
-- Name: Pupil; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."Pupil" (
    "admissionNumber" "text" NOT NULL,
    "firstName" "text" NOT NULL,
    "lastName" "text" NOT NULL,
    "gender" "text" NOT NULL,
    "isActive" boolean DEFAULT true NOT NULL,
    "form" "text"
);


--
-- Name: PupilCode; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."PupilCode" (
    "id" "text" NOT NULL,
    "assignmentId" "text" NOT NULL,
    "groupId" "text" NOT NULL,
    "code" "text"
);


--
-- Name: Role; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."Role" (
    "id" "text" NOT NULL,
    "name" "text" NOT NULL
);


--
-- Name: Subject; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."Subject" (
    "id" "text" NOT NULL,
    "code" "text" NOT NULL,
    "title" "text",
    "studiedComment" "text",
    "commentFormat" "text"
);


--
-- Name: User; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."User" (
    "id" "text" NOT NULL,
    "username" "text" NOT NULL,
    "password" "text" NOT NULL,
    "isActive" boolean DEFAULT true NOT NULL
);


--
-- Name: _ClassToUser; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."_ClassToUser" (
    "A" "text" NOT NULL,
    "B" "text" NOT NULL
);


--
-- Name: _RoleToUser; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."_RoleToUser" (
    "A" "text" NOT NULL,
    "B" "text" NOT NULL
);


--
-- Name: _SubjectToUser; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."_SubjectToUser" (
    "A" "text" NOT NULL,
    "B" "text" NOT NULL
);


--
-- Name: _prisma_migrations; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE "public"."_prisma_migrations" (
    "id" character varying(36) NOT NULL,
    "checksum" character varying(64) NOT NULL,
    "finished_at" timestamp with time zone,
    "migration_name" character varying(255) NOT NULL,
    "logs" "text",
    "rolled_back_at" timestamp with time zone,
    "started_at" timestamp with time zone DEFAULT "now"() NOT NULL,
    "applied_steps_count" integer DEFAULT 0 NOT NULL
);


--
-- Data for Name: AppSetting; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."AppSetting" ("key", "value") FROM stdin;
comment_format_template	<Academic>\n\n<Units Studied>\n\n<SCG>\n\n<Behaviour> <SocialSkills> <Collaboration>\n\n<Overall> 
\.


--
-- Data for Name: Assignment; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."Assignment" ("id", "pupilId", "classId", "eoyLevel", "targetLevel", "actualLevel", "finalComment", "checkNote", "checkStatus", "checkedAt", "checkedById", "linkedData", "aiStage") FROM stdin;
5ac6589191300742d354	1242	6f69f172e874d825216c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3d09804a7bdef90f53c6	1242	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
567a1644ecfc70661ced	1242	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
46d66b07c9dbe8a1f572	1242	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d7fdee35845a1abe95d0	1242	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
46d1bbc2f34dcce140ce	1242	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
860136bc8ff9eec29016	1242	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cc88f374fa08f525e4a0	1242	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
380fc25c5b1c3a2a11fb	1242	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8058d53c35c6ad447665	1242	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f8cdfa3b3031ef461a8a	1242	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fb77518fdc62f196c9c2	1242	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
30e88fd761cd32eb4c20	1242	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
34a86096eef837091c22	1242	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
62bd839c63434cfb559f	1242	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7812139cae19b18c4c72	1242	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7300516a746b7d62d8e6	807	6f69f172e874d825216c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1c8e9babf3bb592bfdce	807	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
09fd672887f1beb450ea	807	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2c539360abfdee7154cc	807	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b10307f5361c1f44d365	807	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6e17af3953683f8702cf	807	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7041a5bd09661a579466	807	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6859251c58c9de26f100	807	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e82fa2782f326ada8ca0	807	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
41d21f8119c65dceed5d	807	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
81e2b311a58a27bbfc59	807	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cbc74c0bb6fe431fbe74	807	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b68c300ab5992064f80c	807	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6c9e89227d2a236d4a3f	807	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ed762b8677584f51779e	807	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
68d1b783ab7968642f80	807	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
765130f160abb73c6525	2287	6f69f172e874d825216c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fe30e052998585541d66	2287	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
194dfbd722316391e295	2287	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
25797958bc00a5b20362	2287	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
078d2444329cb77d44bf	2287	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
79ce685f5e51e4555d82	2287	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
651b0dc6d79d9dc423fa	2287	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1bbd817a9bafbd253a00	2287	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c4b0888f6ccdcec8497c	2287	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c23855411430328e6513	2287	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
80838e2c778457756bca	2287	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d2cd4c415a82ebccdb9f	2287	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d2e208eebb8723686848	2287	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8d9be4ac69a46397fe2b	2287	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
25f767eb4102de3d056a	2287	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
604587ecb252d3a497ba	2287	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3b5790fa3292b26d32d7	2299	40e35d563990a350b589	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9bae64b793d2ce9b2086	2299	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
63597bd509d2cc8f0b94	2299	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eaa58968554c6e0ebc76	2299	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cc163e0d621d9bb1136e	2299	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7727abfe9d116d91cc4a	2299	94b4fdf386152f607a13	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8490b118824f5dcbaa9b	2299	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
225a018b99073024e66a	2299	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
466903c74d8533cbac9b	2299	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9f82a3681d1bc6fd670b	2299	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8155b69297133b0bf82d	2299	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a62bd2b6a26a91a6e5b1	2299	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4684c001268a2957cc64	2299	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
42cdf1d5745b5f327859	2299	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5102c9bfbe85f7dc3f46	2299	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f63962a3828f8c9ab678	2299	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
351cd2f59f2db2cf69c4	2587	40e35d563990a350b589	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fff4bbdeeb155fd5d50d	2587	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d023d464dafa2b6f1932	2587	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d0acb8af3e9e017f6494	2587	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
36b36015ab17ec1b0469	2587	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
32eb6f272fa2e1c9750b	2587	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fc9493a8f2a2d312492b	2587	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3fa919da7b147f72747e	2587	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0e938ef854b832d622e5	2587	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6c865f45d502134db3d4	2587	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
29334d79a4906cace7c4	2587	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b1fa77ecee308c1c91ed	2587	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
88ede585c171612cf23d	2587	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0a5025b912cf9c2dd521	2587	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
73f36e45029a53831f9a	2587	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b362cf2fc06bdddece2f	2587	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a950fbb654777ca08b1d	1341	6f69f172e874d825216c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ffd6573aa025c8c22979	1341	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
195f4afc028ef709e97b	1341	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
72438bfbf64cf133198a	1341	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
67a1528d1a3351b388d6	1341	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
01a9f1f312d20b271743	1341	94b4fdf386152f607a13	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
604cce82a2eebf49fddf	1341	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a56d81044ea18574f090	1341	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
17512b93bde5fcecaea9	1341	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ab64a5214451ed39fddf	1341	89788763a94d40f933fc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2ef0bb579cbef0c4cc45	1341	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dd47a027f1aa68ef9239	1341	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e2e4f83501a9ec88228d	1341	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6b9de60ca3115306bc1b	1341	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3bed506ef2a337b6807b	1341	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
48a301ddc74d940d573b	1341	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f7017a407ef85f023d31	2560	6f69f172e874d825216c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3a0f6cc42c21ea3e8b8c	2560	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9d1d910dbe0fa229b7a0	2560	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b2102f15dcab1868cf7e	2560	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
54dc5ef68bbf0ad4f984	2560	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3ac74cf4ab8040916557	2560	94b4fdf386152f607a13	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e28225710dcd5012947e	2560	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
22670199da6ce6c7d070	2560	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2473a25e50e6ef52e8cc	2560	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
187fe4a8d9129d4ea815	2560	89788763a94d40f933fc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6f0cd6f9a03c094a7135	2560	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a4d3fa3942c3dbacc8c0	2560	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ec394d726de1ca9ac1ba	2560	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
02bcd90901ff7798ace2	2560	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5eabb57290f2df62491a	2560	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1f0facecffcf2924d1fd	2560	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
131776baf2a872c29e5f	811	6f69f172e874d825216c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
58a7e0571f76a2c23138	811	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d51f953bfc06088acb08	811	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c8b742cee4712470dad0	811	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5a7a4986f9346805630a	811	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c1463acdbcc9aabb4ad7	811	94b4fdf386152f607a13	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
daf0d8f3c2ab7b74cb5f	811	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4a47103ec3261ff978f9	811	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
12640f6ac61e6b1a3f6b	811	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a2c33e115ae21b45c4af	811	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
357d2eb071ec4a05f9b4	811	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0967c3f04cd9fdc552d3	811	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
73b095f243227a0394b8	811	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
febb591ca5bf187fba60	811	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2dcc1de02f17099e0377	811	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
877996576b7177c8c031	811	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1355c603dd1ec167d5d2	2413	40e35d563990a350b589	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c2c9a9cfa081d8440610	2413	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c4f47d9417e713f7288a	2413	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
594ef1a7f92245953c4d	2413	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9475bbb731ba9b7c6bfd	2413	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
23106eb0947623838d4f	2413	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
01af4875d16d1cefdc42	2413	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
52b160d4f41a6515a095	2413	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
09d0233c9ddea950fd46	2413	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5c8c55261994f02159f2	2413	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7fd8a82a711fbf7dfd88	2413	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
19b3cdf4f166dfdb3481	2413	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e113d67386c73861b88c	2413	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a6731e3cb50241bbe987	2413	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3705dffcedc386d5f651	2413	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5c6d8bd393a10a6f4467	2413	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
67d8ca33d72a2ae0f47d	1954	6f69f172e874d825216c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
43da4d6fca8ad8761716	1954	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e87addf4f92e4e1b1240	1954	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
080fae26a79181294f6a	1954	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
786026a56e20d2ed3567	1954	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
16f4d7ffb7a79efc5f14	1954	94b4fdf386152f607a13	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e4e19d69a4d5d8bdf202	1954	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
329436af55442fb75d79	1954	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b5ed490f90142abca2b8	1954	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
710e57f5dd6f63d5d20b	1954	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fef2b609c56bbf9ac1b8	1954	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
13e7f319d93128af4ad3	1954	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0fe2421aecc83f9902d8	1954	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e031376d76fa8187c8c6	1954	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
58ac0f01f3232ca60dd1	1954	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
928a1bc0cc695a6b75dd	1954	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
685f91d46629687ce1c3	2381	40e35d563990a350b589	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
19465980590edd6310d6	2381	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8a52930aebfa30075699	2381	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
88dd832dec5ef2d6d3f3	2381	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8a61829752c414226242	2381	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0529b54e743f5b264aee	2381	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
214ce7f0ab8d97fb58e2	2381	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8ffabe1803a713f089b5	2381	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
228577d11ddcdc2d4173	2381	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
db1938cf55961a2f73f8	2381	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d085dd056af2e27b8df5	2381	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
74b0b94d08e2d53fc7d2	2381	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4c982293168d208666a3	2381	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bee79386b0f9b9714bda	2381	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0107f2f5f6e82b5b808d	2381	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
95e86b0606b7d9517b12	2381	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f9581cde19bddb681c84	2416	6f69f172e874d825216c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bf5739bd4939fb5250c2	2416	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
85de9f26162757d1e93b	2416	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aa3cabcf14c8590c87e3	2416	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0b957ee26ff1be709774	2416	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
82d0fbb1bdda018cba5b	2416	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d5d7d3956b9a82b78f7b	2416	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6f11f66d1fa65c7dd5a1	2416	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9cf84953c32fadb19aea	2416	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
30baa88eb10f98a2565d	2416	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
71e0ee62f8490cde7c8b	2416	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3d9ab4f77cbdaa04de9d	2416	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5dc58e15744e8a4d75c8	2416	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
69a2d043612db115a3ea	2416	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
80e465f16723b6d3bbb5	2416	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d2496698e1a1dd91b39d	2416	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a1bd514739b637429c06	2036	40e35d563990a350b589	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
887a7aa644b5b34bd4e6	2036	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5301b8f0b76b198de5ce	2036	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9013110a0412421c133a	2036	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
99dda0e7e3ae86313af4	2036	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c0e07bf3aeebc2f3825e	2036	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
203333275f084f08d1d9	2036	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ba0343cc8a889c40004b	2036	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7eed9a0413dc8b5ed2ee	2036	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c03bc80168b1d9623a9b	2036	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fda320df7544cd25d515	2036	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6df19bcc912d4dc159fc	2036	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8cbf97e86091199f0221	2036	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8393ccaf7d85984da57a	2036	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8161f2d8e6f0134d1e4a	2036	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
acf829614e653de0194d	2036	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
86e89586d063acb5eb6b	2259	40e35d563990a350b589	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
531c418e756d37217b2d	2259	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ec9a70b6512183ad5a24	2259	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
95dcc0b106210eb301a8	2259	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
04495c8cbcf0c349e950	2259	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
22e7086a55695a1566b5	2259	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7ae8fcc9d41624e52ee8	2259	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ed7aacdccf965e3abf20	2259	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
69fbfe775db983c73d5d	2259	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c2e33ff8f7bfde86e293	2259	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c3da9dbf347623426c26	2259	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
31ed4b41732d70f022bb	2259	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c153cc22c35a21474397	2259	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9f9e82af3f29d33ec3da	2259	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e60ca527e79a65291e33	2259	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
32da8b9d12f19545a6ab	2259	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
03da4f249d7cdb0ee401	2254	40e35d563990a350b589	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f665bd1f4014a95b1ef1	2254	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
641ff937c1a9fba3265c	2254	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d6e56d423e418378ccda	2254	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f54e782abf7ba5b020a2	2254	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
62ac520a8831f23edb3f	2254	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e8b1b5d1cf066c6e1fb3	2254	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c6ca0ea5cd71053013fa	2254	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1a43103c09b0a0cbf5ea	2254	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
53291c00b457267b8834	2254	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
957831bcdde72a6a3aa6	2254	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
899e02f7fe175e211bd9	2254	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
90a3859c745388405d2e	2254	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e1b043eebd835d606eb9	2254	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
33868b62b2c24e5b58c2	2254	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4c664e809abafde853ee	2254	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dc174168916e6bc46228	2331	6f69f172e874d825216c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bd6e004f3752a5a835fc	2331	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
53266d94d506a89dbf15	2331	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2266f8c4b38dc4027e14	2331	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8f3b19c330007fc5f767	2331	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
48a1b6a6db8801b056df	2331	94b4fdf386152f607a13	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
430ff775b4309a23c618	2331	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
70e14d8e8b5f80b4ebdd	2331	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ef1ee5bc91c275895669	2331	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ba135ceb8e68dcfcb8ca	2331	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0e10ac265a7862c7946a	2331	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c94975306c03b0a6aacd	2331	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
82344992f7867b56f0a0	2331	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
24d23cafe17347f2c8e8	2331	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2d10f7e63b2a9109d26b	2331	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
439a2fca2e6650cde7b9	2331	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
63116f1ae9b83dcfa0c0	2232	6f69f172e874d825216c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7868fea4dd5ff272bb77	2232	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9dfd88c569b00e947196	2232	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8f3105487985e2a9f132	2232	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
329b061864fd9ba36541	2232	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d22513569df934e53627	2232	94b4fdf386152f607a13	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d0912564c39b0d8a5b87	2232	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c9b53b622f3e822c6563	2232	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
900f2d4d8ce782498f00	2232	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
37cabffcdb76ec3f126a	2232	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d0ee80cf7bf4f7740e88	2232	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5402607341f2c327cb8f	2232	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
33b1a3fa69f18050b084	2232	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4dec030665ed4db5a211	2232	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
df550dbf65287b7e1473	2232	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1993e1fd1669f45bb440	2232	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
36bb7a78d8b307e88b0e	844	6f69f172e874d825216c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d5bfbc710b417438c356	844	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
503c1c8472fbbbe510ff	844	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
40c89e25d035b53382d8	844	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
92dc03d8e8f6e3adf7e7	844	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b5067db3395494a6ee53	844	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ef1a73802606d5204a08	844	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
55cafd2806338c4ef961	844	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
30cb3aa68f71de138131	844	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
874b7778a77be1482c99	844	89788763a94d40f933fc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
61d87c6d5bf1ea1032db	844	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a3ac9ed31a2fd24614ee	844	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4b34666fbc8ed1220bb5	844	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9d7ef2def8f8ca6d58ab	844	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4e48d173f24f0ff20c0c	844	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fda2672fe16c0304854e	844	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8e7db90a8de3e66f39d0	2195	40e35d563990a350b589	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee79071cc5e13545671a	2195	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3af1c67402fb6639cea2	2195	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c7593302b0718697bb04	2195	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5becc669b1bafb2011c7	2195	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
52c117ce8fc4bb233725	2195	94b4fdf386152f607a13	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3473f086a2e2eac1fa2b	2195	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5950d52b0591fbea2fbc	2195	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
542baf486b6282a7ad66	2195	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
04694f138337b02cad4c	2195	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
75a36d7b9a2ae22875a3	2195	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
465215eac3c711609da4	2195	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e586140ac745c7cfa9d2	2195	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3d52b030beb502baf791	2195	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f714db5682c3fc8f8eb9	2195	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3b6992252055150ac54d	2195	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ad216052f9954b0981b3	2135	6f69f172e874d825216c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4c75e00a6f3d71e77271	2135	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
923bdd05781c6595fe4b	2135	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ef6611326ad2dfffe2d0	2135	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1cd453425867db879f21	2135	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c799b0141e9c29395cb3	2135	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
78d47f90609f88888715	2135	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
717913652b9ec704dfa4	2135	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4f4d68f23526dc7b4537	2135	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8d397f83cafd4eafcb77	2135	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5711f6be38e767e7d483	2135	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1e16437d0d4e417e3d23	2135	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
123014232e47f564ff59	2135	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5b17e26749f7a32e8515	2135	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
051ebf77a6708e3be084	2135	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3e711241dff023403caa	2135	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f2d16572367ea5ad3857	847	6f69f172e874d825216c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3eef10e71274dba9ba96	847	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cf206540eb959a549c1c	847	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2b6a9c4c2bd6e1bd8fd9	847	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b1a9bb192dd6ab1af34a	847	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fcca3e0b291d1cfc832a	847	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6f1614b95baf657770ff	847	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
23e9a58239b0fee9e3a8	847	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
91719092576c30291436	847	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
200fe524a15bac0653b9	847	89788763a94d40f933fc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fd089cf951807eaf1060	847	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f311fd4fd664e45a8625	847	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0a328fca9cf05735607f	847	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
14afd7c256e5874da644	847	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
90b2bb7e4ce6cca8e558	847	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
24197817d6f7349b2f94	847	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
32297313a77bf996ec74	1964	40e35d563990a350b589	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6461ad6246d0f0e9814e	1964	ade303ded2be972c6a5d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bdb1cd717767e6d12c68	1964	d13ffb488e82f76b02d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6e31c50596ee8c89da68	1964	ce89e455bbdc905dd759	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3817355b4a922fc3d361	1964	37467e477610598c1474	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4463d0508cf706a69f4e	1964	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
210cc7b43bbe4208a5cb	1964	5e6582e3fe3b905468a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
53e6610abeb003ed34c4	1964	98559a16949b4230e009	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b3608475aec9d3eaa904	1964	6b755b3363797e6feca4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5e8a026c29b763c56ae2	1964	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ea4d7104253a30de9266	1964	310c982ca10f2d32ccb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
64b77333ac8f27dd1422	1964	d421b774241434ccd836	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bd22616bf896d0c5f546	1964	b6bbbaae703037f495be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e9f210eb06af8555e3c5	1964	39b1ab8e2df8c0182786	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a3bdcbf1dd9ba7e14c6a	1964	b630f79c61b93bb9a72b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bc3e6adfffb94042f865	1964	3649e9e9359e028f0892	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f0979cda818b43c3605c	1240	0f69da4484eb0767a720	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e03e3bc3989a20a1c879	1240	0e8355bb2b56b6ccda79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
06bf89d6a1af66388b69	1240	c69b3760d1f706da6487	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6496cc89b6a930649e7b	1240	dd84dd43e18cd13b299e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cd5d58639036da091498	1240	96ead37b6d717ecf4c9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2bc9209b2a9041ea356f	1240	94b4fdf386152f607a13	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
15bf21879cfac469d497	1240	7f4d00c651d1bc22f95e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
89025c18ce0ca3a24826	1240	be3719af94743bcb4e1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
00c635880b74de03eef0	1240	c34e64cafe735b50ed2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2946bd042eadbb817a85	1240	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4228fa9c33417447dddb	1240	f82c00f0fb8ba676d79a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
974f2099607510a3ecc7	1240	d5cde185f3020a8d2c79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
23853093b8f0a52019cb	1240	56061cc582e767291018	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2360a3af27e36c04101a	1240	da2af99abb1da9ed1063	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
78861e61cfb632e2d1bb	1240	ac8714c2edc7a6940ce7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b4fb97b7692fd74a50dc	1240	335b5ae685bab3b50f30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0dc2c1eac1cb4bce486f	809	0f69da4484eb0767a720	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
da25a19e6718c1d5f5ab	809	0e8355bb2b56b6ccda79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9b19a3959bd766f78b27	809	c69b3760d1f706da6487	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ba39a41ac0de12ee6716	809	dd84dd43e18cd13b299e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b60d5bd176f1aafaed87	809	96ead37b6d717ecf4c9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d3d9f93d2fcefd2d2f21	809	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6a331a0ff4dd63c852d2	809	7f4d00c651d1bc22f95e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6775ab48a2ba7c70490e	809	be3719af94743bcb4e1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
22c3e132b7b88343270c	809	c34e64cafe735b50ed2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6eca906081ae6ca541f5	809	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a1130827f3642247a32a	809	f82c00f0fb8ba676d79a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
86563fe4abc38de3498f	809	d5cde185f3020a8d2c79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cc8a420685a6e3192d86	809	56061cc582e767291018	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d73c30438700f38a30bc	809	da2af99abb1da9ed1063	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
430be50b71a438bd9629	809	ac8714c2edc7a6940ce7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
83727cb1178c52f578cc	809	335b5ae685bab3b50f30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ada5c0e31a1f86b5845a	1996	dc667f3c1235dee04b67	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7bf3d047a9c6570f47b5	1996	0e8355bb2b56b6ccda79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3d62334aa41029823fc1	1996	c69b3760d1f706da6487	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7fe9b7ccb59aea1a8f61	1996	dd84dd43e18cd13b299e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee0b209ca12b9c9796ed	1996	96ead37b6d717ecf4c9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c14c50e881a776fac862	1996	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a4f923f565c2d919bb30	1996	7f4d00c651d1bc22f95e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d83dc6371505ad45021d	1996	be3719af94743bcb4e1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cabf5e3f54df3824f0db	1996	c34e64cafe735b50ed2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b0f824c6c597643261ee	1996	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4ca93275f76db9642a13	1996	f82c00f0fb8ba676d79a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d132dc4b3c82fbb9f5f1	1996	d5cde185f3020a8d2c79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4b2c89bfc8be65c6bdcc	1996	56061cc582e767291018	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ac5281c564e34d6a96ed	1996	da2af99abb1da9ed1063	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
748616b979a6746978df	1996	ac8714c2edc7a6940ce7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a213dd47eb88e34203b2	1996	335b5ae685bab3b50f30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0819e16378fc0e96e586	810	0f69da4484eb0767a720	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4ced5033294ae1340018	810	0e8355bb2b56b6ccda79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
20175343cda1613fc321	810	c69b3760d1f706da6487	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cd24076f3f583575cd19	810	dd84dd43e18cd13b299e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b552173466e6cb80ae21	810	96ead37b6d717ecf4c9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
13a254e3728b4e34d552	810	94b4fdf386152f607a13	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
afcc0f32870eb3a57d09	810	7f4d00c651d1bc22f95e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
36b3b94578c0748b37cf	810	be3719af94743bcb4e1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1f5fe9d85ae8b22a52ca	810	c34e64cafe735b50ed2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0f92d706a9597b98fcc5	810	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
755853d9236801291945	810	f82c00f0fb8ba676d79a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c2c4f3867e74d39098cf	810	d5cde185f3020a8d2c79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d5d005f3654e52a29af0	810	56061cc582e767291018	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4327aba711679157a41f	810	da2af99abb1da9ed1063	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6d4392e761d6d9b9daa7	810	ac8714c2edc7a6940ce7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
96c0e6c7cedd6e172fa4	810	335b5ae685bab3b50f30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
24f55fd3e814ccfcab2b	2042	dc667f3c1235dee04b67	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c6c2e38859a42d7fb38f	2042	0e8355bb2b56b6ccda79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f2e8546468f418932bb1	2042	c69b3760d1f706da6487	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e68bf98cd3cb0343cf56	2042	dd84dd43e18cd13b299e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
79c0787f0e5b1c3e4575	2042	96ead37b6d717ecf4c9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
22875bf052fa210beb8e	2042	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f46859925aa5c5ac245a	2042	7f4d00c651d1bc22f95e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d2e4995b0d9eae3bf495	2042	be3719af94743bcb4e1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
630484f928c776845cbe	2042	c34e64cafe735b50ed2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
79c4bffcacca60470eab	2042	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2510b863bc1766f9844c	2042	f82c00f0fb8ba676d79a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4f5fc32cf140d58c1dc7	2042	d5cde185f3020a8d2c79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
62a877f0f0031062f482	2042	56061cc582e767291018	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
669c20676b9760edf951	2042	da2af99abb1da9ed1063	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
55909f0b25edada15727	2042	ac8714c2edc7a6940ce7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
155020dd914aab2f76dd	2042	335b5ae685bab3b50f30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
93dd02686a25ae456c16	812	dc667f3c1235dee04b67	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f1d3e0bfe085564d517c	812	0e8355bb2b56b6ccda79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
61f48f8e6b8c350f31ce	812	c69b3760d1f706da6487	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d1c3662f6f7720c2af51	812	dd84dd43e18cd13b299e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c8465b9ba921fdc58389	812	96ead37b6d717ecf4c9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7f1c890d037bbedf1d9a	812	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
47087483464d468052a2	812	7f4d00c651d1bc22f95e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6c8e8d85cb7fdae41d33	812	be3719af94743bcb4e1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
62612ed3cc5dc09801ad	812	c34e64cafe735b50ed2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
39ffbe4bd304da5508df	812	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8866467d71ce9690bfe6	812	f82c00f0fb8ba676d79a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4035c45fd43345eabbd1	812	d5cde185f3020a8d2c79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
77d8b99aec4557daab94	812	56061cc582e767291018	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3c42fefc426ecb917c00	812	da2af99abb1da9ed1063	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9b08930ca25300486104	812	ac8714c2edc7a6940ce7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d6e464303f6a9d8c0099	812	335b5ae685bab3b50f30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1c553e0fdae7504ac2ae	1800	0f69da4484eb0767a720	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c71f47be1e23599c54da	1800	0e8355bb2b56b6ccda79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9b73c76afdec51f8b707	1800	c69b3760d1f706da6487	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
57ffb3ee569700feafe4	1800	dd84dd43e18cd13b299e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0e9d4bb8dae8334bdb5b	1800	96ead37b6d717ecf4c9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
21b79aacefb6ffd6f8c4	1800	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c9d0d18e777746a3a29c	1800	7f4d00c651d1bc22f95e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8d5daee3088136d0624e	1800	be3719af94743bcb4e1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c86f030dc21f74782c50	1800	c34e64cafe735b50ed2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c5a94751156e29cdb7c1	1800	89788763a94d40f933fc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0178720b1e38de435d11	1800	f82c00f0fb8ba676d79a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b25536f5e58bc3eb2dcb	1800	d5cde185f3020a8d2c79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fe2cd527f794e6ca2c2f	1800	56061cc582e767291018	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
609ba9f31f65714c895e	1800	da2af99abb1da9ed1063	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
95a133e2d1c6f72131fe	1800	ac8714c2edc7a6940ce7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9f32798356764fa109dd	1800	335b5ae685bab3b50f30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
76061b67c39b6deb5b6e	2296	dc667f3c1235dee04b67	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1099de3fd011df3f2b28	2296	0e8355bb2b56b6ccda79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
66c41875ff6be37fb2eb	2296	c69b3760d1f706da6487	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
756d30a14d86179f925e	2296	dd84dd43e18cd13b299e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
50ba12e2b1180bdb840b	2296	96ead37b6d717ecf4c9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
809fbaa3c1b6a22902de	2296	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
051eef9d5b67c630ae8c	2296	7f4d00c651d1bc22f95e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
70add64041935883e783	2296	be3719af94743bcb4e1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f1233441c9063bc43005	2296	c34e64cafe735b50ed2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
10ce36dbb33fa81473f5	2296	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7932c8036885717a77a0	2296	f82c00f0fb8ba676d79a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3befa5cb113d74173aeb	2296	d5cde185f3020a8d2c79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dc215d25330ef23e2d9d	2296	56061cc582e767291018	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b2212aa1541acdabc728	2296	da2af99abb1da9ed1063	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3e4819aada0a95c73ba8	2296	ac8714c2edc7a6940ce7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
772f6194b7d40575b886	2296	335b5ae685bab3b50f30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7d364e0ca0a18116dbe0	1526	dc667f3c1235dee04b67	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4380f5fd3afff56f507b	1526	0e8355bb2b56b6ccda79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4886a02f9b5b391c4a06	1526	c69b3760d1f706da6487	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
08921465abe62f2efc6c	1526	dd84dd43e18cd13b299e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
41b83f7750f9680d0a06	1526	96ead37b6d717ecf4c9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8525abd5393b6fee074d	1526	94b4fdf386152f607a13	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aabbaaf228cc12d2a1af	1526	7f4d00c651d1bc22f95e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8bd35f322690c60ea31f	1526	be3719af94743bcb4e1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
de3f4793e7c989b631a0	1526	c34e64cafe735b50ed2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
61fcd398760955968cea	1526	89788763a94d40f933fc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d0e20c4f5f37451a2456	1526	f82c00f0fb8ba676d79a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
59e15bc8b13550459448	1526	d5cde185f3020a8d2c79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b695fe77d9340c0d1a61	1526	56061cc582e767291018	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8739757dd19df29941d5	1526	da2af99abb1da9ed1063	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6848cbd57dc1be0bd058	1526	ac8714c2edc7a6940ce7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1cb2297d76a126c097bc	1526	335b5ae685bab3b50f30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2d59c724297e5192a018	2561	dc667f3c1235dee04b67	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
976431f0f9af0d4243e9	2561	0e8355bb2b56b6ccda79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
410d8b894d2054f2bdb4	2561	c69b3760d1f706da6487	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fcd6d30f98d23b55b1d7	2561	dd84dd43e18cd13b299e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
92bc870463939e3705c9	2561	96ead37b6d717ecf4c9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
309858afcb4ac2f53317	2561	94b4fdf386152f607a13	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cc8ac5e31170475114e9	2561	7f4d00c651d1bc22f95e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bc35a17c857e0eee00d8	2561	be3719af94743bcb4e1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f8b9349e231b79fd7242	2561	c34e64cafe735b50ed2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
12d67ce551e06942ebe9	2561	89788763a94d40f933fc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f06d3bee2ef773b31edf	2561	f82c00f0fb8ba676d79a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1c97d221fd3bc4a17a31	2561	d5cde185f3020a8d2c79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5fbf2ca9a2f59ca14c01	2561	56061cc582e767291018	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8884a5f2e43608abf9dc	2561	da2af99abb1da9ed1063	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f5a628b6fdf8842acc1b	2561	ac8714c2edc7a6940ce7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d02e293d0224d2f2a447	2561	335b5ae685bab3b50f30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c59bc22f6095621ae146	1554	0f69da4484eb0767a720	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
db9422b6243a03d37833	1554	0e8355bb2b56b6ccda79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1a3554b5c82e2c24b784	1554	c69b3760d1f706da6487	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b39c39d72e64813dc2f6	1554	dd84dd43e18cd13b299e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f634176ddd65cda02be5	1554	96ead37b6d717ecf4c9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e04f58df39c6d2404daa	1554	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
38bae12556942938cb09	1554	7f4d00c651d1bc22f95e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
51885d8244c961af968d	1554	be3719af94743bcb4e1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4a7a089a95b6da05850d	1554	c34e64cafe735b50ed2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c8696c814f46610e68b4	1554	89788763a94d40f933fc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9d142ebed405d14d6e80	1554	f82c00f0fb8ba676d79a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
519db1e8b9319f21b37b	1554	d5cde185f3020a8d2c79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f77a4847818c4f836ec0	1554	56061cc582e767291018	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c650d8d99ca79f9a0365	1554	da2af99abb1da9ed1063	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7a5765122cf4f6d69935	1554	ac8714c2edc7a6940ce7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bdb78f27658af18a5471	1554	335b5ae685bab3b50f30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
db9acb7d447a03908d0c	2592	0f69da4484eb0767a720	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b9a970fb186f28fbc8c8	2592	0e8355bb2b56b6ccda79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ec46645830d87c17792e	2592	c69b3760d1f706da6487	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
96ac7095bd9cacbdbd7d	2592	dd84dd43e18cd13b299e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9d4ce3b3c02883b21007	2592	96ead37b6d717ecf4c9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4e72b7bf05c0c55f72f5	2592	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
118689c04df08f2bd0f9	2592	7f4d00c651d1bc22f95e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0431487de2b34f720a96	2592	be3719af94743bcb4e1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
55e5603277c4738dd981	2592	c34e64cafe735b50ed2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9657f094cda6f59a38c8	2592	89788763a94d40f933fc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ff8da1161f6b018e7c56	2592	f82c00f0fb8ba676d79a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b2b02c970fa103552d25	2592	d5cde185f3020a8d2c79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6cb9cb53c8ede04cac8e	2592	56061cc582e767291018	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0338a0b2a39754874956	2592	da2af99abb1da9ed1063	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8feeb74741488f56d64e	2592	ac8714c2edc7a6940ce7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b1514705f66a1eec258f	2592	335b5ae685bab3b50f30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e32bd18be6f4c0d16afe	2422	0f69da4484eb0767a720	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bf104996f9b438ac686c	2422	0e8355bb2b56b6ccda79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d790eb3764f454964e9c	2422	c69b3760d1f706da6487	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f559782ea694961c6744	2422	dd84dd43e18cd13b299e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5c8e27b1cba9a270b04e	2422	96ead37b6d717ecf4c9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
85d3be5cae47cc0e5309	2422	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1c434e79591ae535d39b	2422	7f4d00c651d1bc22f95e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3024249bbb3e2215e29d	2422	be3719af94743bcb4e1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5af6be42a849de4db0ec	2422	c34e64cafe735b50ed2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b5ae68aec80dd2287398	2422	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fd34674b169ff3f520be	2422	f82c00f0fb8ba676d79a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
422d4ad2c2072c0d276e	2422	d5cde185f3020a8d2c79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4cd27bff060d208d7671	2422	56061cc582e767291018	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b87a20c8283c4de833be	2422	da2af99abb1da9ed1063	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8d858e51b51335847617	2422	ac8714c2edc7a6940ce7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7a21199b2d38cfe01243	2422	335b5ae685bab3b50f30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d41e3d53b168998e1b97	2562	0f69da4484eb0767a720	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a6b47f698038fc06c282	2562	0e8355bb2b56b6ccda79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9f4e9a03b5ee982f32a8	2562	c69b3760d1f706da6487	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dca59170b884c78772ec	2562	dd84dd43e18cd13b299e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aa9c019658503fe1ad54	2562	96ead37b6d717ecf4c9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cef98787f137f6cee044	2562	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
07b25c8a7a9a8c5b8576	2562	7f4d00c651d1bc22f95e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f2e4bd387b4a09fdab90	2562	be3719af94743bcb4e1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1b4c466232abdbb640ec	2562	c34e64cafe735b50ed2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5dc79c542f4038fb635f	2562	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d136a51671cfdd457973	2562	f82c00f0fb8ba676d79a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
21811a6dd3a0e70989d0	2562	d5cde185f3020a8d2c79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
94b73f84b50b8c6ab1f7	2562	56061cc582e767291018	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c86ee07f0f12d198065b	2562	da2af99abb1da9ed1063	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1bc18db96a466cdf136e	2562	ac8714c2edc7a6940ce7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
feb2fa29bafdea9479dd	2562	335b5ae685bab3b50f30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1fd29813175cdcae555b	819	dc667f3c1235dee04b67	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9eb27c97bd8a4eef14c9	819	0e8355bb2b56b6ccda79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6d577ef51b8230d21657	819	c69b3760d1f706da6487	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
976e69fc687a0e1d0ad1	819	dd84dd43e18cd13b299e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5f148db54f4a185bcbbd	819	96ead37b6d717ecf4c9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a0e488999e73bd2621fa	819	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6b64950375a2650703a3	819	7f4d00c651d1bc22f95e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bf0d76ca8bb17096a528	819	be3719af94743bcb4e1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
05a334c1dc24c583db90	819	c34e64cafe735b50ed2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
737e9652f12719d79742	819	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1bd0a089b159f8f26cd7	819	f82c00f0fb8ba676d79a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f784ddf5e6d016976903	819	d5cde185f3020a8d2c79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5d6154bf2447fd3f6bcc	819	56061cc582e767291018	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d6bc4359b4df63df8d93	819	da2af99abb1da9ed1063	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e4c1792907937bc7c144	819	ac8714c2edc7a6940ce7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
311c840f0057538ea357	819	335b5ae685bab3b50f30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a5cacc8f5ac6bc2b4464	1244	dc667f3c1235dee04b67	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
28f7ebc4c5f631441ef8	1244	0e8355bb2b56b6ccda79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
85b1bdf64f143159bac3	1244	c69b3760d1f706da6487	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3d4e2829acfdca249802	1244	dd84dd43e18cd13b299e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
538b8618ba8ef21ac976	1244	96ead37b6d717ecf4c9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2bee949c56476e767de4	1244	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a3436d6fe38c1e2e4f46	1244	7f4d00c651d1bc22f95e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f74a4f46c41e571a7f2f	1244	be3719af94743bcb4e1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c8389937e21390883c84	1244	c34e64cafe735b50ed2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9109ed280afde082ed48	1244	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
59f45503cc0a34a757b5	1244	f82c00f0fb8ba676d79a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0903ef7712320b2c08c2	1244	d5cde185f3020a8d2c79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
439372b873f6828940d4	1244	56061cc582e767291018	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
318740b3b84247b3a21f	1244	da2af99abb1da9ed1063	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
505528f99375744100fa	1244	ac8714c2edc7a6940ce7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
62eebaeb697d5fa0e7bc	1244	335b5ae685bab3b50f30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2d4c606e8114bf1a8ec4	2207	dc667f3c1235dee04b67	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2cd36b4fb26914e93110	2207	0e8355bb2b56b6ccda79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
af11cdabd98233be466f	2207	c69b3760d1f706da6487	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ac79f5d7d42d64941ba3	2207	dd84dd43e18cd13b299e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
700bda09cc3edb8fdab7	2207	96ead37b6d717ecf4c9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
38b4e2b4e73c5c51ef96	2207	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
85110d615914a35c85e6	2207	7f4d00c651d1bc22f95e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6d178316bf3481f43d81	2207	be3719af94743bcb4e1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
348f5a12c1aecfc327da	2207	c34e64cafe735b50ed2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
894d821206db160d163e	2207	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
90e0d9a52578d6c4c71c	2207	f82c00f0fb8ba676d79a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8b302072abbc27345305	2207	d5cde185f3020a8d2c79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0e5f60137d9b4bf1b104	2207	56061cc582e767291018	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
40bc174ea3b14bd8f6e5	2207	da2af99abb1da9ed1063	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3dc4d4db3025ef6cd690	2207	ac8714c2edc7a6940ce7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9daac0cdb5be9a8571eb	2207	335b5ae685bab3b50f30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d9a2fab435e3847b2103	2356	dc667f3c1235dee04b67	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b6aa98d86215efc34850	2356	0e8355bb2b56b6ccda79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b65146076bb9d5c1a54f	2356	c69b3760d1f706da6487	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
02cc4616ecc30cffbbde	2356	dd84dd43e18cd13b299e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
63d37ea6daf37dd05423	2356	96ead37b6d717ecf4c9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a56ca743a85f3898786e	2356	94b4fdf386152f607a13	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
293212aa8c1575a0a30d	2356	7f4d00c651d1bc22f95e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9150ae7170f49c83f1e3	2356	be3719af94743bcb4e1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6c6593ee89cedd4c7c54	2356	c34e64cafe735b50ed2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2dbe6ec0bee8e2fd9cc5	2356	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c3e843fd97762f471292	2356	f82c00f0fb8ba676d79a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e7a2e9d79f2759cc9902	2356	d5cde185f3020a8d2c79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fc0f47b898e25cf8b422	2356	56061cc582e767291018	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e038930b4c9a278d89f3	2356	da2af99abb1da9ed1063	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
08283ff04c60e491c6cb	2356	ac8714c2edc7a6940ce7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
be0b232fd1063f3b93b6	2356	335b5ae685bab3b50f30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6c8148a63d91cc215176	1057	dc667f3c1235dee04b67	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e53b308b6c6f0cee0f8b	1057	0e8355bb2b56b6ccda79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
39af221c0b71962ff80c	1057	c69b3760d1f706da6487	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ae1bff1aecf1b7e06e2d	1057	dd84dd43e18cd13b299e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
12e2a38eaa2c4ac79e53	1057	96ead37b6d717ecf4c9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
08c728975c9ce8a97810	1057	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
862ad77d799233834754	1057	7f4d00c651d1bc22f95e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e637faeba80e4f1be11f	1057	be3719af94743bcb4e1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
424eeb3827a103301b72	1057	c34e64cafe735b50ed2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
431e43cfd279f70480dc	1057	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e359668d6c0d3efd7321	1057	f82c00f0fb8ba676d79a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9390d89e6196129afd41	1057	d5cde185f3020a8d2c79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0b6817f453ee6ee99bca	1057	56061cc582e767291018	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4f23e66c4267233c49a0	1057	da2af99abb1da9ed1063	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4b10b0a670c52c87c728	1057	ac8714c2edc7a6940ce7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8c1625d2ae6e4609ae50	1057	335b5ae685bab3b50f30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8ac0a4f669d28fffae79	1689	0ab34b12507bd5ae0386	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
28c4e3f703e3044729d4	1689	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ead3a640ec8a981521f1	1689	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8c71f84295e4581be3a0	1689	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bd7ff3458a054b3e7269	1689	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
578fe1b565632f72538a	1689	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fa037bc8bb6e0d904b6b	1689	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c0c9f9e75eda037b2bf9	1689	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
deb323bd229d9b2f5aed	1689	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
734b158784947f0ba884	1689	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
594048c60e6f0c3c41e7	1689	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
937c5eb488f5efe6cbcc	1689	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
60b48c703962a878929a	1689	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bf1390108285181d4a09	1689	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
545ebac9360f89c6e90d	1689	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
67a87cfb8ef303162ede	1689	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
412345e81f6419b18b0b	1690	0ab34b12507bd5ae0386	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8ef29bbc569609324750	1690	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cdfd579e1344a9f33608	1690	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dba2531c7951aac63039	1690	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c0d138e1e7d7d826681d	1690	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1c29ca7178eb2058d66d	1690	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6074b4a5af16c292b93b	1690	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f0f5297a582dcdebd50c	1690	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
48b2eaab6dfa7b415e4e	1690	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dbf5ad61fcf736e0310d	1690	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ab87d7adde5e6a2bf6f4	1690	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8dcba35bfd8c99a2eeb7	1690	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
31ad39249fee538fb502	1690	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a425d01474360c49a2e2	1690	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6459004cf2261837d042	1690	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2129acfe0de6fac2e632	1690	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6a7c96ba02c7e9eba672	1053	0ab34b12507bd5ae0386	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
db9de560692a085e49ce	1053	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cbfbf7b0e143b9d40edb	1053	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8ba927e3f09361e105c0	1053	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
92b8e3bdff4fede17bb6	1053	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ba535888a2ec465bfa10	1053	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e92895d1e819cd31878f	1053	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
547582e0cb1395cd251c	1053	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
81abf816da951080328d	1053	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
575fa2b3eace3fe90c2f	1053	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7e34b1b12771c137039a	1053	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cf5a4ec0a304c77ea116	1053	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
639e49eb86581203c7c4	1053	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
403e70ab1a9daff0b03f	1053	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ea22a3003d9723825139	1053	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e2947f509fb6cc669d4a	1053	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e4d4137588c0853a54f5	2405	09c52f1ac992e352a953	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1ecd60b757584ee91892	2405	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d5ec84f921699da919f4	2405	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d97a47b38f4f91a6f0a1	2405	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9a0d96382a18c1f166ec	2405	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fe00e01cf81318f7ad0f	2405	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b3a117fe9a6fdcf81128	2405	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a21e6c3163ddd3a2646a	2405	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
85e68298110b42b4aec4	2405	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a0bfcf10c7d4d3b50b67	2405	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b8328dfd2330a81ad456	2405	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ae81dd07c6d52bdea9fb	2405	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8b989a9f11ed2d9aee38	2405	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a0f18f372bbddd669f26	2405	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f426430bf9e2e33d2ccc	2405	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fb724ccdab8e8eadf998	2405	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1880663c389df7a931de	2257	09c52f1ac992e352a953	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
42d8cab148509dc304d5	2257	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e9bb88ba2c77a1e10e9e	2257	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5f96aa7361a7ccdfecac	2257	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4cad9222345b20e93929	2257	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7524cdc54329a5d8852a	2257	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d6c8f4ba5573611ce9a9	2257	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
24d1c23f7ea46c3c730f	2257	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a7be1755fda5c6737496	2257	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e9101155ede3f7be2c65	2257	89788763a94d40f933fc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1f7c58fc288300f9c9d1	2257	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3c3a45dd50963998346a	2257	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5fbe8d18351402761a65	2257	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5d82f7f9d103d0c82434	2257	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ad5e09b6db2598e49db2	2257	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4f2513fb80fe4ce3a911	2257	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
964ee15ad9360c55e026	2312	09c52f1ac992e352a953	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e9b29b6059ea64dc7dfa	2312	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
82affd7141ac71eb0d40	2312	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c11f955c9aa605b0bd60	2312	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
da726bbcea9bd1378b59	2312	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
829d0cb9ab7dd1741834	2312	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5b3af8c51542c2e1f279	2312	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ff9501ced96c282fd675	2312	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ccf749a35dd998489023	2312	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
64bd3fc69929029c9693	2312	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0dce0c0b9b5cc1ce46d1	2312	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6a11193f34044d827e26	2312	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9027343778c06e68912f	2312	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ad4658d4d4af64085ba4	2312	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
78651de20c879715998e	2312	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2f2b8a89fd749f070e96	2312	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
142eb7ac198d774488ce	2411	0ab34b12507bd5ae0386	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4972738c0329e2224c79	2411	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7432f948972e3e0bd840	2411	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
12e99133913fe83301b9	2411	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d2ac7e6aac35e446e42f	2411	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a6622dd5d1801836a5e8	2411	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
826161bf176bb818ce4c	2411	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cdce830034a678667f47	2411	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dafb6c092b08b05428b2	2411	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f8ac6a9ab746c0b9dbf5	2411	89788763a94d40f933fc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c6b0d0f3a04315aca123	2411	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
644b390fc909a1390fa2	2411	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ad0162766fd693da6f00	2411	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7ba7772781f8b13999bc	2411	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4b03b1d42490363fb9a6	2411	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5e46ddd76b337fdbb970	2411	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cc71096e976038de6f2d	1071	0ab34b12507bd5ae0386	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3e6b29bff4944866d268	1071	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dc5b38baf4442c945959	1071	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aa0f6c6356ea12aa6da3	1071	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
18bd54ef54768dc41622	1071	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
654713a2b967653fdc27	1071	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8cab5a21ebe9fb300a6a	1071	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ac61aef7e8c0960485b7	1071	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
078fe33e567bcc74f67a	1071	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5ab8bbe04581a80de20b	1071	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a413b1120bdcc8ac3418	1071	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
945db6a5fa4a49286cb0	1071	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e4503bbb13d893fdb16c	1071	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aa11f200e55e1c5ada43	1071	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
830ed0ba2db28ee6b766	1071	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
326fd0e5d57dc1325f9f	1071	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b1ee74a45da31e97333b	1955	09c52f1ac992e352a953	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c3f578b6a8a19bd08661	1955	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7724b8255e066d2e50de	1955	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b6b4924ec143dc755999	1955	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0af71c6534af12b6a1d7	1955	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0e41173289af32678a66	1955	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c0fac57d2fe8dbf4b40a	1955	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3ae4c0ac054a0a118948	1955	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e90832320539ef5c028b	1955	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9db32532dff2ef733c52	1955	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e70e1cedb926dc6ef8e8	1955	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d3510f4a69e8989fc85b	1955	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
23e70661346b002fb9f5	1955	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0bdee6bcb0730b50f242	1955	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
33aab0ebea74c7ebd59a	1955	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cc5f3ed0fd59cb52b60d	1955	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e125a40abff13c1bcc9f	1804	09c52f1ac992e352a953	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6a07da510b9cc9ac9706	1804	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
42d79c212cdba9cb8477	1804	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f19234e3f4e8bd4f13ec	1804	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d4cbe79fcd2fd036e48c	1804	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c0fd950aa573adf55335	1804	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cec0e05e6123878ae452	1804	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
71e1b282d9aef48d2b6d	1804	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d7d396d3811df8fdc0a2	1804	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5f0e396f817f6b695fde	1804	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0f7de25183a1ed8508a4	1804	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
75b6b974932e716ee25a	1804	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
49b455b79b2fbd03b524	1804	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1fdaebe314e72d9d6ccc	1804	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b58e1da6c6acaff94e57	1804	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
262371278551656b91fc	1804	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
af8006d948069bdfdeb7	820	09c52f1ac992e352a953	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3beb00674b22e3ad252f	820	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
791860a833de94355097	820	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bc2d2d2848618e5d8c43	820	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8e5534bbf102d89bd2e5	820	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a998b3b90c17078db86d	820	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d3fd5e445f36da38cc61	820	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4d3014c70d234280fdcf	820	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c0fca0851957f53b0b1e	820	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f857fbacaa416700636a	820	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b3dd3269973a97c6f417	820	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5bcb88763d0945d3c62f	820	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e5804faf7905de2a4510	820	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2dd308cdfef52f3eec62	820	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b90375745ce2973b1533	820	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
96a84cb236e9cf161ad0	820	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b047d79fdf641ae50046	2334	09c52f1ac992e352a953	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7b4d1847391eadb9544f	2334	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
787438c10d95e6513df7	2334	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0ee221b190a831d59430	2334	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
05d11a27dad0880d6e2b	2334	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0050f1ff59a41b2b4ea5	2334	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4bce694634f0a7ac089e	2334	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
77ad85589ac891f5f7f5	2334	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e78390203bbeeb0dcd90	2334	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
74f2101e18a18997d63f	2334	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4b10dd45b5180cfeb262	2334	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3723a5a30fd7dbc161fb	2334	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7287116a5362481c7a9f	2334	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
94842fbf0b7ae04f2dc7	2334	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
946b979b49ed00545ad0	2334	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
90e5af461ac05895a8ad	2334	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d7fe637fe2a147c8563d	1389	0ab34b12507bd5ae0386	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
45a3965a18efc864664c	1389	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
88f19bc70ffd6373c706	1389	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8ae3c77ff07bfaf22b3f	1389	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e790c88aeb79e9a32a61	1389	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
df2cefe31f83dad68944	1389	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
228ceea580d208d309b1	1389	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f3a3ec54d79e2e46fde5	1389	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fac69b450e0803e61df4	1389	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7244897b486369adab0b	1389	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8fb604e4f96bb6369910	1389	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5c1c3dbf4da1904656c5	1389	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
67033ddbcd233411e337	1389	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
18c3dd33ac0cb465b427	1389	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a4d776b81792d4db5509	1389	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3f2c71c2d77dcaa2c9f9	1389	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
611c70c26a4bd1ca9599	1571	09c52f1ac992e352a953	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c0affb1676ef542d51d7	1571	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d1733d64ad98fdc4ee42	1571	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d2c749b6a39b957aecf9	1571	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
92cb84c0beb85b0eefb9	1571	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
741b920962919b7d63ab	1571	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a5c514e6a1950995d514	1571	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
83999b387adaad5485a4	1571	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3662f5ce90d4363cd17b	1571	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7d80362e212e717deaed	1571	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3391400ec58f20dbebb7	1571	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5d1c1b6d7300044d04a4	1571	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3e5ea1c8e7cc31994d6a	1571	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
db7c0a9bebad05ec9e87	1571	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c7ffc32912b7814ff533	1571	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
83812832d547614405de	1571	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
50f27b83662c5ffef808	1117	09c52f1ac992e352a953	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ba0b781822017d4a974a	1117	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d4b5af9fabbfeb1d4ccd	1117	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0c3f61954323b5f8c7af	1117	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b28699561db148a57650	1117	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1e38e05284e0b7f29803	1117	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6f4fa6b6ae5d3e43336c	1117	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ab615101ff6854b9c16f	1117	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
468e4c9b8750cfe2753d	1117	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
547c96b37ddcc895c8fa	1117	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e800700ffc4463cce025	1117	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2d5b733c572cf3cbc6d2	1117	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0f4e478d99b17502f60a	1117	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
31141f8ba92d016739e0	1117	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6f61050f7cdf015da1a9	1117	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
83357fd649544f154348	1117	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a0f52656818b97dcb3a8	2566	0ab34b12507bd5ae0386	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8fcd5abc9aa55cb3b3da	2566	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e730cc8f991f699e8e28	2566	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b9cc7ed080c31015e3bc	2566	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fc69033b9a5ff422e678	2566	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2eed27222d8bc768e9e4	2566	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bd95da0650c7d595f07d	2566	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1b87a3e6f4b2dff79b7b	2566	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bf2e458beaa163f04ac6	2566	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8bb1a78aa9f3b89926a7	2566	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
96062e32d067d2c9c738	2566	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2cf04fc5213c96c5a35a	2566	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d04279c80af52746aa7c	2566	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
465a9ebbea273cd4248c	2566	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
42ec7bfe94ae9a519223	2566	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5e2fc44ee1fda9e939e3	2566	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d20ad78068030db95fb5	1594	0ab34b12507bd5ae0386	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3cebf4b0cd01df048062	1594	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2e6b3e510255bd7a67d8	1594	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
596b6b5b9865a108a354	1594	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
da7cf99cb9fe48bf7bcc	1594	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b98975c760ce09ae3152	1594	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
64d66f55e81afd5ff3d1	1594	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
41ba1c0e50b1e9bb2a4c	1594	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0fe6f16b30bb4504ed10	1594	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b34c6cf70f2c1d112bda	1594	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6d2113dde22abb3998db	1594	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a0ccbc58dece9835d5f7	1594	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fc3ccf987b0af664d199	1594	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3956420676540693ea25	1594	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3ec583414c3bf1ac5234	1594	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
20e2f2222cfbbc1c8c13	1594	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7a0760098245ae9fbeb8	1593	0ab34b12507bd5ae0386	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ff5532d27447a4ac5161	1593	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1ac60e63e83641cb0962	1593	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
01fdfdbdfe9dfb0e6658	1593	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b0b2f24e515d78b02dea	1593	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f9b14988ce38e84e4b8b	1593	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b93e0bdaebf025b941f9	1593	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ab2b0aad17a114a96ec7	1593	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dbf84a72690807ece3f6	1593	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
79c53d757fcbc509154f	1593	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eb4a4e84b39a2c5750b7	1593	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ed406ccd56b2ab8adb6e	1593	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e6f82c8b4f14d41b8a36	1593	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5a2b12be1d0cfb0e0b3b	1593	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2649e025566aca33ed71	1593	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eb32a63319c7dc110e21	1593	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
72e7957b91440b0c70cf	2199	09c52f1ac992e352a953	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f9eb2e982175d8536eb9	2199	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
18310c0a2f57f96eef38	2199	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3def523cb1eee87faf1c	2199	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
18f9261a91b644316f6c	2199	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3290ac76e900f2015aff	2199	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
15ea51c07d2156127a3c	2199	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1693119f1e226b0d9ab5	2199	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ea61cbdddb803f43ff87	2199	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f3f1b638ef8ba3e0eda7	2199	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dada0361cfda564c4f9d	2199	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5a69e0da484debd15fe9	2199	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e02b1adf03c6933aecbd	2199	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
be7dfff9bef197d90be8	2199	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9081b5ac16e6f71d5fdf	2199	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6829b47fba3a1d46b1dc	2199	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5acb2a4c17c0f3280383	826	09c52f1ac992e352a953	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fb90c3958576f28809bf	826	d15d7b9fd2ed7e46dd53	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ec16c0e687c0eb44bde4	826	2a9db3e19b93ce9c355d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d67bb9e429ce88bcc777	826	87ccbe2ce9a680ff93a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b9891a9c075b55328684	826	859e00b64f3807defcfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
26e0b4c8eeec48efcc32	826	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2c14cc2876926fdca021	826	a477427b675c489d36f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d81ebf58a8445355ca13	826	0e634788a37d82abbef4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
524758281c4e32487577	826	d7cef190d3651254a780	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
68d2c563c23a96390fd9	826	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
30797a5cf8c634960c64	826	98d9b41d2eb6d4fc61c8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9de0cd692ee90e9f2f2f	826	dbe35887465186683743	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
440a3465cd5360610ace	826	abe9d4482e5fa2c5fa0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
59145586dbebdb721756	826	7e1c155a4ca800a30d86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
be7e67c078639514d92a	826	8cdd523b41f1a026f4de	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0e6d7c2c05bcd57d2de9	826	061ddff82e7b822b7ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f4d39cf75c5f5b1722b5	1688	05504bdd46170dd3f4db	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d9b873834910c7a7dbd2	1688	6db953d26eebb36848ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
14063cc92bc570873fd7	1688	b4ccb16d04bc3b14647d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ea690dd575bc432c7699	1688	a8ff0afc8a2eaa57d68b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9d34368fbc5bdbb6639f	1688	22eceb505d094f53c53e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ca7439bd4de4f6a2bf46	1688	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
456a332aabaece62c006	1688	273f5b438cc3414ba44f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c7453ec6dff1598bf94f	1688	ad8ed2ff59798b5df7ba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bf95ffa8b2c7e7339fee	1688	8639aeda2bf449728d9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e77f08a832cfa9c0671c	1688	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee5d805fe2e62cbf84f4	1688	b4f371ad7826b1be4b8f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e4b8bf3896ca1c0a049d	1688	bc02a71806dababec353	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
be3f8eba83cc6de4d53d	1688	0af02ffbcd58d7253c76	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
255246361d4e579066f5	1688	4d66ba7137d37d690550	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b7a554e178a9f00c2b48	1688	d9c04b00fffcf9b34420	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ac5214c65681c79b2068	1688	011da44732c7f182de66	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
31e1e07012b3f0d324c5	2131	05504bdd46170dd3f4db	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c252cb0c4ee594d245be	2131	6db953d26eebb36848ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0c89082be29bc179abf9	2131	b4ccb16d04bc3b14647d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
add8657b64bbd3379fd5	2131	a8ff0afc8a2eaa57d68b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a64730eb6b4e768e4bc9	2131	22eceb505d094f53c53e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
69196cb4d0d6e45424ee	2131	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1c177504d036c856f1cf	2131	273f5b438cc3414ba44f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e2d28752344ea2e8e1d9	2131	ad8ed2ff59798b5df7ba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8bb61709c4ae8a1a253e	2131	8639aeda2bf449728d9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c5642b76e8638f9c5f1a	2131	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5a38986eb6354acf8d3e	2131	b4f371ad7826b1be4b8f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
56bf830af73492ab9c53	2131	bc02a71806dababec353	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
25a5f8156e7a76c86916	2131	0af02ffbcd58d7253c76	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9597557a427a12b82f5e	2131	4d66ba7137d37d690550	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f17b4849fb06a2e3dcf0	2131	d9c04b00fffcf9b34420	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
02b3e330d76d28f525b3	2131	011da44732c7f182de66	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7f2d43d07e154bd5bcdf	926	05504bdd46170dd3f4db	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eb33d861d4d7c4db7487	926	6db953d26eebb36848ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5c6f8ef7c09bd028a2a2	926	b4ccb16d04bc3b14647d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
69e366bffc6dbaabb79d	926	a8ff0afc8a2eaa57d68b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ca57e7be6f380ba7cc65	926	22eceb505d094f53c53e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
315acce8076cd1d19d3d	926	94b4fdf386152f607a13	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6b6208b5befe515e2e98	926	273f5b438cc3414ba44f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1b2027ff7f287794b22a	926	ad8ed2ff59798b5df7ba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
56015e726ac4de759cd4	926	8639aeda2bf449728d9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7d260b17a6481f09bdff	926	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e3faa9226dae43a05513	926	b4f371ad7826b1be4b8f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b340c3d0dac752e10fe5	926	bc02a71806dababec353	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e62aaf4da93803ef4a7e	926	0af02ffbcd58d7253c76	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
282f2add00149f597d78	926	4d66ba7137d37d690550	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a9e34ee1cf59207e38fa	926	d9c04b00fffcf9b34420	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
959c084a1af1e28088af	926	011da44732c7f182de66	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
586996bc8c42eee4e615	2291	83e93fed391058692f49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
18cbea4fb3d899e16093	2291	6db953d26eebb36848ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee6d9dd149ab79e6daa4	2291	b4ccb16d04bc3b14647d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ddd2cb2d16e074642650	2291	a8ff0afc8a2eaa57d68b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
63dbdb278de0f3b2a92e	2291	22eceb505d094f53c53e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e7789374e923e7561d9f	2291	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
959b79d16ba30adf37a9	2291	273f5b438cc3414ba44f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
896aefba7109eef54079	2291	ad8ed2ff59798b5df7ba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dd8a719289bcf120cd63	2291	8639aeda2bf449728d9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
61e6bbda57ef08edaec5	2291	89788763a94d40f933fc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7faae4032483db836f1d	2291	b4f371ad7826b1be4b8f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d7b422097cf23223e9db	2291	bc02a71806dababec353	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0c8c75310df90d463c75	2291	0af02ffbcd58d7253c76	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a5c64c5117d4cf20ab4a	2291	4d66ba7137d37d690550	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b0ee7a367fc462c8b56d	2291	d9c04b00fffcf9b34420	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0dd9621605b5e6deec1d	2291	011da44732c7f182de66	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c7d53c00627710b5c6b5	831	05504bdd46170dd3f4db	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ec976c380aadfe4beba4	831	6db953d26eebb36848ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
80320cdc76919a476d61	831	b4ccb16d04bc3b14647d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
14727f3ab958e7f8c429	831	a8ff0afc8a2eaa57d68b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3db7653af6f073492f7d	831	22eceb505d094f53c53e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4ee226178e8b9fa332f0	831	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7c96aee170ff817149bf	831	273f5b438cc3414ba44f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0e3648dc8f55cc69f017	831	ad8ed2ff59798b5df7ba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
27c7f1049dbf7f5158bb	831	8639aeda2bf449728d9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
889ef5771e4317ce9b26	831	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
45b1e6ff9f617e2a4091	831	b4f371ad7826b1be4b8f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6b5ccd24bfc02fa1af48	831	bc02a71806dababec353	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
64166590b89c7675d2c1	831	0af02ffbcd58d7253c76	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c5621d9bbf68acb8496d	831	4d66ba7137d37d690550	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
192a842bc8751d629530	831	d9c04b00fffcf9b34420	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e1e66a005cf1bf39597d	831	011da44732c7f182de66	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0355fbd8c02c611eb918	827	05504bdd46170dd3f4db	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4b5afdae3691f9a474d8	827	6db953d26eebb36848ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
24e2f014694aaf171de9	827	b4ccb16d04bc3b14647d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a0f477240a9d94948f4e	827	a8ff0afc8a2eaa57d68b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
768063cb980fd0ddcb63	827	22eceb505d094f53c53e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b980fbd0b2461f845c5e	827	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
981489d41eebd7362727	827	273f5b438cc3414ba44f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dcbd09cffca29aadd865	827	ad8ed2ff59798b5df7ba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
37680fc4377aa8f62c8f	827	8639aeda2bf449728d9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ba68dcb8aa391a75b036	827	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2cfb069785fa3888b488	827	b4f371ad7826b1be4b8f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0b3b2c4241a40c5ce8f9	827	bc02a71806dababec353	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b03790bb1fe5c31887dc	827	0af02ffbcd58d7253c76	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cfb1abb92d25cfd4b672	827	4d66ba7137d37d690550	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
daa579bda0ed90e9159a	827	d9c04b00fffcf9b34420	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c5e7c648ed51b5431245	827	011da44732c7f182de66	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4747f2c84097ac9a752a	1246	83e93fed391058692f49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4c587cffde08d3e7eb7d	1246	6db953d26eebb36848ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
db29e470fc47336fa02a	1246	b4ccb16d04bc3b14647d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
751a97db48b85fb5d2b5	1246	a8ff0afc8a2eaa57d68b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
85755d3493c284382efd	1246	22eceb505d094f53c53e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0e9e12b9cc54ebde1d4e	1246	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4ae626f8d7987d62802b	1246	273f5b438cc3414ba44f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
39866bae3c341f556b98	1246	ad8ed2ff59798b5df7ba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fdb5df13fcfd8ef06793	1246	8639aeda2bf449728d9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7e580e6ac564b05aa1c0	1246	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
deda7235b6d623b6cbf1	1246	b4f371ad7826b1be4b8f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bc31d59c03b799bc4277	1246	bc02a71806dababec353	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c8aca990bd90ddb0f4fd	1246	0af02ffbcd58d7253c76	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f024d0714ccea2c3dd20	1246	4d66ba7137d37d690550	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a3af68bac5ca45088afe	1246	d9c04b00fffcf9b34420	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b26493064fa2ecfe63d9	1246	011da44732c7f182de66	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
73b45937b1c2183ce228	1247	05504bdd46170dd3f4db	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6e2735cb3ed3bbb45e82	1247	6db953d26eebb36848ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
79a5ceb04a0041d76d86	1247	b4ccb16d04bc3b14647d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e980507bd9c2abbd76b2	1247	a8ff0afc8a2eaa57d68b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1318abf0ec8e06853f45	1247	22eceb505d094f53c53e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2c013b5c79bda6b71af1	1247	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0f4191eae9517a027d5d	1247	273f5b438cc3414ba44f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c4d45911adda741112e7	1247	ad8ed2ff59798b5df7ba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e5fbba4d97f9d9f1562e	1247	8639aeda2bf449728d9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
026a4073066bce7dbe3c	1247	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
505d0db427ec5d5395ee	1247	b4f371ad7826b1be4b8f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
36f40f32ae9a89d13671	1247	bc02a71806dababec353	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
529a3ca1a26615ab9a1d	1247	0af02ffbcd58d7253c76	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a42231bfb8b419265335	1247	4d66ba7137d37d690550	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
717c75e24d83c19f94d5	1247	d9c04b00fffcf9b34420	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5a3953789c7987e17e73	1247	011da44732c7f182de66	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f1b3b0ecef773af94e98	1531	05504bdd46170dd3f4db	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
abfd00b4be68688adc03	1531	6db953d26eebb36848ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cbaecb6479aae417c75d	1531	b4ccb16d04bc3b14647d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3c82f41ff5b5f51bcd76	1531	a8ff0afc8a2eaa57d68b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d80fb323b376669a88dc	1531	22eceb505d094f53c53e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b4b5762e1db8a72cb8b3	1531	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
081a9ddbe3b227b8a778	1531	273f5b438cc3414ba44f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
532c94690375cf7b147e	1531	ad8ed2ff59798b5df7ba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9a187f5f1ffe77d49efe	1531	8639aeda2bf449728d9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f919e640b0e0b7a3a7bc	1531	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4b429b0f3adb90472455	1531	b4f371ad7826b1be4b8f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7d6685a6a2f21268d6cc	1531	bc02a71806dababec353	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
86179da0891ca6173776	1531	0af02ffbcd58d7253c76	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0870d5f5addd613da6c3	1531	4d66ba7137d37d690550	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f6c97df2eb5cc97cbe32	1531	d9c04b00fffcf9b34420	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2103cf49e81a81e23007	1531	011da44732c7f182de66	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
845b9f169c59a657b645	2033	83e93fed391058692f49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
87a02fc171df8104e95a	2033	6db953d26eebb36848ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bc73a9b2a5eb3037a2d0	2033	b4ccb16d04bc3b14647d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dfea017a6bff1d545ea5	2033	a8ff0afc8a2eaa57d68b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f988eec093924a773212	2033	22eceb505d094f53c53e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fc5993cfcdc54d5ac66a	2033	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ac511729cc3acfd50c43	2033	273f5b438cc3414ba44f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c4680de3d8db8ec5f562	2033	ad8ed2ff59798b5df7ba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a41bb179c95baa08cabf	2033	8639aeda2bf449728d9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
98b94ef7351d32e07f30	2033	89788763a94d40f933fc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
761b07bd72d204e1d4b8	2033	b4f371ad7826b1be4b8f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1486e07c0b5f9b1a235d	2033	bc02a71806dababec353	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
428eca4afa874808d48b	2033	0af02ffbcd58d7253c76	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e61b2f01da1eacc6a12a	2033	4d66ba7137d37d690550	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c96bbd4eacfb941d5f15	2033	d9c04b00fffcf9b34420	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
94d9fa7ba9dab2a1af8c	2033	011da44732c7f182de66	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
41c1e370c9ea0e8b2514	2252	83e93fed391058692f49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
43dbf03b2e53568d72c5	2252	6db953d26eebb36848ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c83af5b9cbb53ec25940	2252	b4ccb16d04bc3b14647d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
92da00abf89a93638e83	2252	a8ff0afc8a2eaa57d68b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4b39099346ef724c4d22	2252	22eceb505d094f53c53e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f61dd0b6b69c152f4dc2	2252	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2bc94936949e0849d603	2252	273f5b438cc3414ba44f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2f09344061316b3ef02d	2252	ad8ed2ff59798b5df7ba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d0d7f6e936c8422bf907	2252	8639aeda2bf449728d9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bc8433a382f47bbbc3a2	2252	89788763a94d40f933fc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b2f92bac94204bae4d88	2252	b4f371ad7826b1be4b8f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fb55730eb40162e80282	2252	bc02a71806dababec353	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4da0203c49d45db177bf	2252	0af02ffbcd58d7253c76	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
832b51e989e4a620d69c	2252	4d66ba7137d37d690550	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8a148fbc5531399c7f7d	2252	d9c04b00fffcf9b34420	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9d2d8634b0a027a06d88	2252	011da44732c7f182de66	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
022b9955312a0ab57f55	2563	83e93fed391058692f49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
504d2a9ef542a0946c46	2563	6db953d26eebb36848ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cfcf6edbfeacec580b89	2563	b4ccb16d04bc3b14647d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
46defd6942149b6ef209	2563	a8ff0afc8a2eaa57d68b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8d0ce5f466e6e1903be6	2563	22eceb505d094f53c53e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
553c46c69d7423e53232	2563	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6f75c68152c9e024a5cf	2563	273f5b438cc3414ba44f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
956893384294ffe396c7	2563	ad8ed2ff59798b5df7ba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3bbf9a06b5e2d9f3775b	2563	8639aeda2bf449728d9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee34665b373e18866cda	2563	89788763a94d40f933fc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f2155bda1385809ca515	2563	b4f371ad7826b1be4b8f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6d027d772214f491d25a	2563	bc02a71806dababec353	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f76ecab963a3f37b6ff5	2563	0af02ffbcd58d7253c76	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b696d4f1d75fa01d30b9	2563	4d66ba7137d37d690550	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b8a8f5ac004cbc21ad58	2563	d9c04b00fffcf9b34420	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c305749a8e2270fe3f9f	2563	011da44732c7f182de66	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ce3accebecf636a47929	1897	05504bdd46170dd3f4db	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7ef4aaa7d5d4e84ba3c6	1897	6db953d26eebb36848ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
557f914d3f16fc1cf07b	1897	b4ccb16d04bc3b14647d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
591b0373e9c135718b09	1897	a8ff0afc8a2eaa57d68b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1689348bffd6f047be81	1897	22eceb505d094f53c53e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9143c6562839744e8044	1897	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2321e73093ec5138c55c	1897	273f5b438cc3414ba44f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ef9affc9029aee67a946	1897	ad8ed2ff59798b5df7ba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2455384df98352a9b76c	1897	8639aeda2bf449728d9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4e803e182743e3da7d6b	1897	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0c25b7799abc78888592	1897	b4f371ad7826b1be4b8f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b6aadfa7949e4087361d	1897	bc02a71806dababec353	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e790afab54334d8610fd	1897	0af02ffbcd58d7253c76	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4d05792764b2207505c4	1897	4d66ba7137d37d690550	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d20a89db3e6bda040f37	1897	d9c04b00fffcf9b34420	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
171f220c460a2bc33115	1897	011da44732c7f182de66	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
32a7458f8a2ce77c8e2b	2133	83e93fed391058692f49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9ff15e4770935279d9f7	2133	6db953d26eebb36848ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0d723e20ce2ad4479be5	2133	b4ccb16d04bc3b14647d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a13d4dec084b3f0f2601	2133	a8ff0afc8a2eaa57d68b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
717940a6c290cfc2248b	2133	22eceb505d094f53c53e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bbf151887aed25ae486c	2133	0aeb1b79309469c06c74	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9dff3f9195f0cd1df831	2133	273f5b438cc3414ba44f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b515a5e3d5f06ebe815c	2133	ad8ed2ff59798b5df7ba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5a159051edf5125cd194	2133	8639aeda2bf449728d9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e413aabb887db355da92	2133	89788763a94d40f933fc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d3f355991708b52ae0f0	2133	b4f371ad7826b1be4b8f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
10fd72a39aed2807851e	2133	bc02a71806dababec353	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
582b4b951be46801c310	2133	0af02ffbcd58d7253c76	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9ddc1deedf32d5d6120d	2133	4d66ba7137d37d690550	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d36dacba96fa50b16ef9	2133	d9c04b00fffcf9b34420	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a45a2827091aa427c950	2133	011da44732c7f182de66	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b8d3225f785d0f30d588	1118	83e93fed391058692f49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
870c1e226ec67d0a439a	1118	6db953d26eebb36848ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3851480118b2aa8c956d	1118	b4ccb16d04bc3b14647d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c29750815b42988c1112	1118	a8ff0afc8a2eaa57d68b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
639b49a8281e133220c8	1118	22eceb505d094f53c53e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
29afb7cc4a0c5a89a337	1118	567b97f8b1de47a17d7d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fce9ddda26847fb97a83	1118	273f5b438cc3414ba44f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
405cb1136f5bd91cfbb4	1118	ad8ed2ff59798b5df7ba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f29cc4d7db9453003a76	1118	8639aeda2bf449728d9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
031baf657935b3a2b015	1118	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
767f5d68cb7f7b102019	1118	b4f371ad7826b1be4b8f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c9b779dc45a5415eef34	1118	bc02a71806dababec353	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ec4c4ec578b8fcfe763f	1118	0af02ffbcd58d7253c76	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6ce19d4b5820132faafd	1118	4d66ba7137d37d690550	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b1efb63e7c5480e56e2c	1118	d9c04b00fffcf9b34420	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
86b4bfa46082e0951fc4	1118	011da44732c7f182de66	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3f3e38a029967f7b354b	2565	83e93fed391058692f49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eaff4aa218cab3f48812	2565	6db953d26eebb36848ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c87fb6af6a6a7595d367	2565	b4ccb16d04bc3b14647d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
47751e713f74d521e37d	2565	a8ff0afc8a2eaa57d68b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
94197202e84c71eaf981	2565	22eceb505d094f53c53e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e263ba90e1f5f87714ca	2565	94b4fdf386152f607a13	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e64bf64cd06bb287c52a	2565	273f5b438cc3414ba44f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5671800e2c9b80716b3c	2565	ad8ed2ff59798b5df7ba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
edbff3ed4bcefaf3a750	2565	8639aeda2bf449728d9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3b68359e92d80a7c799f	2565	0dae9224680dbe8d6673	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ebd719b936722599a6b0	2565	b4f371ad7826b1be4b8f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4ee2ca852e7a04e4fd1c	2565	bc02a71806dababec353	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ef9ead84270fdaef4c8e	2565	0af02ffbcd58d7253c76	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ec897519cfc2cdda5e22	2565	4d66ba7137d37d690550	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c97ae8b6b631577c4d23	2565	d9c04b00fffcf9b34420	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
38d64cc726fe2f177f4d	2565	011da44732c7f182de66	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6004c832a0fb6328cf2f	1367	05504bdd46170dd3f4db	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
16fe0537b630fbdfbecb	1367	6db953d26eebb36848ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9afa3173ea77be9f6ec6	1367	b4ccb16d04bc3b14647d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1ce21422c84bc35080ad	1367	a8ff0afc8a2eaa57d68b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cc297596c5cab808fe35	1367	22eceb505d094f53c53e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
02c8719e95eed2544d69	1367	a7de483e1b81ba1ede04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a1a8987f5adde6df1fa5	1367	273f5b438cc3414ba44f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ae62e2ecbeaabda7d5fe	1367	ad8ed2ff59798b5df7ba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f9dc6698205a65195a7a	1367	8639aeda2bf449728d9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f9907f3c415ad1dac743	1367	492be8b28f568157cbef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
02a7475369702c5d9463	1367	b4f371ad7826b1be4b8f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
46e3952a92698f374ce3	1367	bc02a71806dababec353	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
da8acfda5014272d2986	1367	0af02ffbcd58d7253c76	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fa46e7504c5a6ae8d114	1367	4d66ba7137d37d690550	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0887906de75dbdd5d08e	1367	d9c04b00fffcf9b34420	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7e099ff0729e100cc2bd	1367	011da44732c7f182de66	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3cb47eee1254402d4182	1073	83e93fed391058692f49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
50f468d61265698dd843	1073	6db953d26eebb36848ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4cf9adff2db3e72e189b	1073	b4ccb16d04bc3b14647d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7f0fa1050468a38f399e	1073	a8ff0afc8a2eaa57d68b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1ed177a3cf91455842f9	1073	22eceb505d094f53c53e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c423f71496bcbba5cd69	1073	94b4fdf386152f607a13	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0014601b435176a55e42	1073	273f5b438cc3414ba44f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aaa80073f479c52655be	1073	ad8ed2ff59798b5df7ba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ae97df38b71f367320ba	1073	8639aeda2bf449728d9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ac54f3df6e7232241090	1073	89788763a94d40f933fc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2fcda212ca4d7612bfb9	1073	b4f371ad7826b1be4b8f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c097c3f05d5b7226d6be	1073	bc02a71806dababec353	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c5a0a3a182d6c4d92506	1073	0af02ffbcd58d7253c76	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
46dce7430161adc0fb61	1073	4d66ba7137d37d690550	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5e8ee46bd6ad22ecc60f	1073	d9c04b00fffcf9b34420	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a4c7a7fd07ba4cfaaf54	1073	011da44732c7f182de66	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bffcab58d5512dd596b6	2377	83e93fed391058692f49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a4f0176bca6db2d259b5	2377	6db953d26eebb36848ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
05b9154c6eaeb048b554	2377	b4ccb16d04bc3b14647d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fa3a6613a0189257745e	2377	a8ff0afc8a2eaa57d68b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1ba53955365ac9b23012	2377	22eceb505d094f53c53e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3a0d516391558d5ae2b8	2377	94b4fdf386152f607a13	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b3835bde5302469c8b6e	2377	273f5b438cc3414ba44f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8fc4403346b9d1d5b586	2377	ad8ed2ff59798b5df7ba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
57a9b072409808c9f9e9	2377	8639aeda2bf449728d9d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8076b7ccf8e56670dd06	2377	d3f8ea02cba4d989b5f4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
764743c1aa80818b2c22	2377	b4f371ad7826b1be4b8f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3007c90f23acf152c67f	2377	bc02a71806dababec353	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bc40ea32500c7ba98cf2	2377	0af02ffbcd58d7253c76	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7efee5ae6b5e2d7871ac	2377	4d66ba7137d37d690550	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c0a3306e30b489faedca	2377	d9c04b00fffcf9b34420	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
85b48624b3ac33cbb479	2377	011da44732c7f182de66	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
98e4b70748982226922c	2285	ccd8ac5ce51d687a65b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fbe95570b0c8fad7a8f8	2285	ccdd93e22ce8fde4aefa	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
05b3d7c8179dbd29a346	2285	2ee37f6a5b5afdba3728	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7529742ce7e73807e586	2285	81f8ac02e568df5d9552	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
00a2caacd23af0f587e5	2285	4a894a4948ebc3d652ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
53f1d6da9272281fa760	2285	3ea70d2daf15198c70a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
993c19d9b9f53f82c003	2285	62a6aa3edf4fe2dee2d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3f5ea5da49c4eb95f8a0	2285	cd4dde106f1d53c9aada	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
885f2a029b4bb9bd7a7d	2285	011f0cc85f1ad1cc22e6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
86da995cbb71a77dfe57	2285	be0802e0a01e96e75eab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
16f0bdc8a80060b6792c	2285	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7f0ed2884f2c8f7b9d7f	2285	042f404c813aed6ee038	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9df4191320932e985d81	2285	9538a1857d657c0f004b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8fb34c36cf357e1d4f46	2285	d74ad31e1cd4b3e1a678	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
65b898c7d4265d82e310	2285	998e79a5a2df8f596fa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4cd99ce20b8153ff8c48	2285	bd6f2ca2ed437f9499d2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1cd4367d8eb50f73d77a	1763	15d8cb818f560a27d9a8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
30b09048c13ef62c463f	1763	ccdd93e22ce8fde4aefa	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d8a20be44461e0316017	1763	2ee37f6a5b5afdba3728	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ce44a4b31477a4b280d3	1763	81f8ac02e568df5d9552	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0b01bb94c9ec22d3420c	1763	4a894a4948ebc3d652ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c5b3f51ee79699412c61	1763	3ea70d2daf15198c70a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
569fb3172e6561e55a95	1763	1c95850b2ece593eeb02	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cf25c3408145d5447549	1763	cd4dde106f1d53c9aada	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7c58491dcd89bce67289	1763	011f0cc85f1ad1cc22e6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6fb7ac93388fbcb6c16b	1763	be0802e0a01e96e75eab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
442bd15ebee138b7548e	1763	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1abd03b3698ef61acf71	1763	042f404c813aed6ee038	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3260c0a7bf289aa8c69f	1763	9538a1857d657c0f004b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1de628cdacf74f56753f	1763	d74ad31e1cd4b3e1a678	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f4cc20935822dd4d0b7a	1763	998e79a5a2df8f596fa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c99045b7432dad5b140f	1763	bd6f2ca2ed437f9499d2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1e702814b4ed61790a3f	14	15d8cb818f560a27d9a8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4b1ec6e70818f56d8749	14	ccdd93e22ce8fde4aefa	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
251f3c199694fb1a9061	14	2ee37f6a5b5afdba3728	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
65f07955f9af4bd3d866	14	81f8ac02e568df5d9552	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dff6a41e42daf0150b5b	14	4a894a4948ebc3d652ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
02c0754109dc41dce9b0	14	3ea70d2daf15198c70a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7eb72704449d87f5e14c	14	937d339e9ff57e16d3c9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
26d8726b4151c0126ef8	14	cd4dde106f1d53c9aada	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6e1b118605d4b2c159ab	14	011f0cc85f1ad1cc22e6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4f83c1a0d78932ffe772	14	be0802e0a01e96e75eab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5ea7bc50d8e8d859efd6	14	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9b57693ca98cc1181199	14	042f404c813aed6ee038	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c8157039d73e07c40ce5	14	9538a1857d657c0f004b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d5deb05b174b19284783	14	d74ad31e1cd4b3e1a678	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
06f3440da67ae0292e00	14	998e79a5a2df8f596fa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2249d8b2b47fb8949218	14	bd6f2ca2ed437f9499d2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b38b77474c8e81431ed8	719	15d8cb818f560a27d9a8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6b45e4c939d4e920013d	719	ccdd93e22ce8fde4aefa	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0c7dc5f07e1ea0d3458d	719	2ee37f6a5b5afdba3728	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1b1b3f5d61b6e8200666	719	81f8ac02e568df5d9552	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
508e50db868df9c2fa0b	719	4a894a4948ebc3d652ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a98f6637c0f9bad74938	719	3ea70d2daf15198c70a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
947bfef2520f87ea0972	719	937d339e9ff57e16d3c9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ac1e4300774d7d96ae71	719	cd4dde106f1d53c9aada	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d49d152002091aade457	719	011f0cc85f1ad1cc22e6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
56b1257b2c981e83cad3	719	be0802e0a01e96e75eab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7b0797bd219a3744f75b	719	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0de969c1983b105c2f73	719	042f404c813aed6ee038	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
26fd97e6baaaf3e5ec38	719	9538a1857d657c0f004b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0d52a03ad6a3f0c79d5b	719	d74ad31e1cd4b3e1a678	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d58448297df26094fee9	719	998e79a5a2df8f596fa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
47dbddae714eccb7cdd5	719	bd6f2ca2ed437f9499d2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7eb4aadd9d5651291311	2567	15d8cb818f560a27d9a8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d3af53637c3bfe04c857	2567	ccdd93e22ce8fde4aefa	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8150fa00af5c80dd1a78	2567	2ee37f6a5b5afdba3728	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
87955dd25110c1b67bbd	2567	81f8ac02e568df5d9552	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cf2f51b0181c7e174948	2567	4a894a4948ebc3d652ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0b6b212761366cb6fece	2567	3ea70d2daf15198c70a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9cc1bdc390b9f9ea6c57	2567	937d339e9ff57e16d3c9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9d56957f8a9e660115d1	2567	cd4dde106f1d53c9aada	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee51e1b257069819ed7e	2567	011f0cc85f1ad1cc22e6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d29703cfcde9d1fbd623	2567	be0802e0a01e96e75eab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3885051842cb74b1db3c	2567	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7687a66e60785fac8dc9	2567	042f404c813aed6ee038	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f74d38b82d5970254d03	2567	9538a1857d657c0f004b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d7b58714b1cbcd0a95a6	2567	d74ad31e1cd4b3e1a678	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
016707b665482f20fa63	2567	998e79a5a2df8f596fa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f5529583c02e4b9df62a	2567	bd6f2ca2ed437f9499d2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
755f9a17d50617a87c2f	2568	15d8cb818f560a27d9a8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bd4413fe9b1769e9cdde	2568	ccdd93e22ce8fde4aefa	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
67efb874570996b93df2	2568	2ee37f6a5b5afdba3728	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a61cebe5872a09a43ed8	2568	81f8ac02e568df5d9552	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4ce69e0b038abb8c0084	2568	4a894a4948ebc3d652ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8e95538f00fd27c8a100	2568	3ea70d2daf15198c70a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6bf93da37c5109cacdb9	2568	1c95850b2ece593eeb02	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c6b84a1a33b7c1778d2b	2568	cd4dde106f1d53c9aada	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
971c9a6526936e22dc12	2568	011f0cc85f1ad1cc22e6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
32810311655f96f5a582	2568	be0802e0a01e96e75eab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7aacd874a344dad77b28	2568	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4ff0e28243fd7b320bf6	2568	042f404c813aed6ee038	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
747c9754e2d6dc701b10	2568	9538a1857d657c0f004b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
119e0f3503c28cf77833	2568	d74ad31e1cd4b3e1a678	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7a99d8cd8da42534115b	2568	998e79a5a2df8f596fa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
62d8492410bee8241d3f	2568	bd6f2ca2ed437f9499d2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
316ddf350d11e61eaced	2031	ccd8ac5ce51d687a65b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c82f201d9a7176eb4e36	2031	ccdd93e22ce8fde4aefa	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4f4e78095da7bd89aeb4	2031	2ee37f6a5b5afdba3728	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8e9eb9868ca16436e446	2031	81f8ac02e568df5d9552	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4dfc9dd3c8746370e312	2031	4a894a4948ebc3d652ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
49a5c0d00d3796efcffc	2031	3ea70d2daf15198c70a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9c5ddb3212e4ab68b8c4	2031	937d339e9ff57e16d3c9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a4d2fc966cfd1f11ef38	2031	cd4dde106f1d53c9aada	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dacebefb3e2104994700	2031	011f0cc85f1ad1cc22e6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
577e1e0702db8400c29a	2031	be0802e0a01e96e75eab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f4690e5e3cc71872be5b	2031	67cb84d4c26e5522b07c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f2f595f0eab223446499	2031	042f404c813aed6ee038	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5a5476b87ccdd0eeceb7	2031	9538a1857d657c0f004b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3ebc8381faeb40221752	2031	d74ad31e1cd4b3e1a678	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8ecb44ecf1f8a0ca8ba6	2031	998e79a5a2df8f596fa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
905e2922257791ec9134	2031	bd6f2ca2ed437f9499d2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
76df25282f8ab9f4804a	1997	ccd8ac5ce51d687a65b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b0dd31b64d7514c9ec1c	1997	ccdd93e22ce8fde4aefa	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
64601fd6c9ed21e80be7	1997	2ee37f6a5b5afdba3728	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
17e0c2d746a3d49f4f9e	1997	81f8ac02e568df5d9552	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
10ab371630c5daa40480	1997	4a894a4948ebc3d652ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a3b3de7ec53c10874e20	1997	3ea70d2daf15198c70a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5fe020fbff5e921b452c	1997	937d339e9ff57e16d3c9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9a655377fc8dd176df42	1997	cd4dde106f1d53c9aada	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0d397d2713cf80eb2d2d	1997	011f0cc85f1ad1cc22e6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a05a7e68ea72d573dee2	1997	be0802e0a01e96e75eab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4ac4dd644c981324f88f	1997	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c060b2da2597d8a6bee4	1997	042f404c813aed6ee038	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7dccc56ebaea9b3ac21e	1997	9538a1857d657c0f004b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6ed7b7dc813782a4202e	1997	d74ad31e1cd4b3e1a678	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
48e0b8913adbce598bad	1997	998e79a5a2df8f596fa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3bcc514f81f692b1f822	1997	bd6f2ca2ed437f9499d2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bf5ccd1361dadb0fd13d	1812	15d8cb818f560a27d9a8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5307639f2993f0acaca8	1812	ccdd93e22ce8fde4aefa	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c2d3c5a96079ef6b484e	1812	2ee37f6a5b5afdba3728	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0f7934bb35195edca3aa	1812	81f8ac02e568df5d9552	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
34f8db021e38c546cf97	1812	4a894a4948ebc3d652ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
22590271a9c10ff28872	1812	3ea70d2daf15198c70a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e83f6902b3ec502f1661	1812	1c95850b2ece593eeb02	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e3d78c13023de0cd0fa0	1812	cd4dde106f1d53c9aada	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0453109f28f1354b029b	1812	011f0cc85f1ad1cc22e6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1a36261c30e1a9f24b99	1812	be0802e0a01e96e75eab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9a4a73d02bf36bf2436d	1812	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7957fe569260c789a1ef	1812	042f404c813aed6ee038	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
311eb93965cffc1bc77c	1812	9538a1857d657c0f004b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c378ed69725951d2f9d2	1812	d74ad31e1cd4b3e1a678	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
87c62e9622c317bef43b	1812	998e79a5a2df8f596fa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c8843bcf0972b6af4125	1812	bd6f2ca2ed437f9499d2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0b9e0cd770301bf8c8d3	1901	ccd8ac5ce51d687a65b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
71fe424cfd8f7a8fc635	1901	ccdd93e22ce8fde4aefa	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
53c07c2bf9a88c3f6c4d	1901	2ee37f6a5b5afdba3728	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
70f1ee08f2ec25a5b781	1901	81f8ac02e568df5d9552	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
89751bc86247268fda73	1901	4a894a4948ebc3d652ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
15e70cad2c8af266787b	1901	3ea70d2daf15198c70a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
25fc49bf346a40960b1c	1901	937d339e9ff57e16d3c9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0c54972aab0c54402c5b	1901	cd4dde106f1d53c9aada	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
59780e959cc064f02e82	1901	011f0cc85f1ad1cc22e6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
09e05005b887dcf44f92	1901	be0802e0a01e96e75eab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f4f0060877ab822c37ce	1901	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
efa19808ccce1f82a446	1901	042f404c813aed6ee038	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a50c5526aa3bf649175b	1901	9538a1857d657c0f004b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
12e8f0988a248fbcd0ce	1901	d74ad31e1cd4b3e1a678	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3e625727b5003cab61ce	1901	998e79a5a2df8f596fa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f214d9eb8929f3c97cf2	1901	bd6f2ca2ed437f9499d2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3f6d98944e2654b47a46	1816	15d8cb818f560a27d9a8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2cd66e442fe79fa38c89	1816	ccdd93e22ce8fde4aefa	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c89145b6b368f5a7db9a	1816	2ee37f6a5b5afdba3728	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7160317f3fbdc2225a5b	1816	81f8ac02e568df5d9552	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7fcf67fd8d4a57e62e70	1816	4a894a4948ebc3d652ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7ce3df43f8ac86e07592	1816	3ea70d2daf15198c70a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
029531e8f9855793c357	1816	937d339e9ff57e16d3c9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7f01b7fbe86a73962ff7	1816	cd4dde106f1d53c9aada	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9aa0b0ab6afca792a42c	1816	011f0cc85f1ad1cc22e6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
52bf552276050fd2bc60	1816	be0802e0a01e96e75eab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b08d9f00589d51b08133	1816	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1c9daa749cf9bb7e82f3	1816	042f404c813aed6ee038	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1d20b47885b2d7bdcb3d	1816	9538a1857d657c0f004b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
568e16fb6b3dbc85ee0f	1816	d74ad31e1cd4b3e1a678	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2a0659bf35794dce1f13	1816	998e79a5a2df8f596fa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
648d3acc21aac3384657	1816	bd6f2ca2ed437f9499d2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4751f410567dbbf4866e	1701	ccd8ac5ce51d687a65b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d287f5f4c171c0ac6ab8	1701	ccdd93e22ce8fde4aefa	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7be4e8eb415ef712ccc4	1701	2ee37f6a5b5afdba3728	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9d0ac04091429185c4af	1701	81f8ac02e568df5d9552	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0ba3368a1b35e129f4f6	1701	4a894a4948ebc3d652ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b4dd6a5119a274912032	1701	3ea70d2daf15198c70a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
615d3ebec0cd215c6171	1701	1c95850b2ece593eeb02	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3439dd17e9920e66901f	1701	cd4dde106f1d53c9aada	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ecb7935d057f5ea5d52f	1701	011f0cc85f1ad1cc22e6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
07b85c1c7a7be84ccf34	1701	be0802e0a01e96e75eab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
97ffc10f0dcd9879453a	1701	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e5f1b2d22ae229aebf65	1701	042f404c813aed6ee038	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
89792779ef73d401a63b	1701	9538a1857d657c0f004b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
458a5559652b452b8228	1701	d74ad31e1cd4b3e1a678	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
46456fe133b3901fd4ea	1701	998e79a5a2df8f596fa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3d78d01416ae8bb747c2	1701	bd6f2ca2ed437f9499d2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f9fa2f74c7625f622f7a	2184	ccd8ac5ce51d687a65b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d647d1ee015283e57c11	2184	ccdd93e22ce8fde4aefa	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8c8f414638468d108f8a	2184	2ee37f6a5b5afdba3728	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
da9c85960ef4b989e849	2184	81f8ac02e568df5d9552	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1e10624505ff5d584042	2184	4a894a4948ebc3d652ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c6693874d782f7ea9f09	2184	3ea70d2daf15198c70a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
40a4c3a582b86c05f770	2184	1c95850b2ece593eeb02	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
49434354130d04b9be5e	2184	cd4dde106f1d53c9aada	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f84d67881a87466f9716	2184	011f0cc85f1ad1cc22e6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a1446378f6f7682874be	2184	be0802e0a01e96e75eab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d533d90e2dcc84ba995b	2184	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c6baf164a12da082f697	2184	042f404c813aed6ee038	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6b6b0c401b6bdf72e08c	2184	9538a1857d657c0f004b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5c1aa0300cc49143982e	2184	d74ad31e1cd4b3e1a678	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e3c75fe0de79a0171eaf	2184	998e79a5a2df8f596fa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
60b243967035e04489f7	2184	bd6f2ca2ed437f9499d2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9f681ca9b78c10e83b08	1818	ccd8ac5ce51d687a65b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
995634179568be4d8d77	1818	ccdd93e22ce8fde4aefa	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8eab6c833ef7dd5cda27	1818	2ee37f6a5b5afdba3728	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c420fc71777bf0a9280e	1818	81f8ac02e568df5d9552	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0ed581abc678b8a9b969	1818	4a894a4948ebc3d652ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e0d287a996628a8ea402	1818	3ea70d2daf15198c70a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7372950d77d7dc1497d7	1818	937d339e9ff57e16d3c9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d905c272125b72aed9e5	1818	cd4dde106f1d53c9aada	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fd23229ee924568b2630	1818	011f0cc85f1ad1cc22e6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
12425755383036296714	1818	be0802e0a01e96e75eab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cfbc29ed193b0148d526	1818	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7e3359682748c636c391	1818	042f404c813aed6ee038	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d7539fc206e2a7dd06a1	1818	9538a1857d657c0f004b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c79515ba38b2141bbe99	1818	d74ad31e1cd4b3e1a678	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b5c4c1f0df79efc1394e	1818	998e79a5a2df8f596fa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d5ff25b72bf8df41b496	1818	bd6f2ca2ed437f9499d2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6142e31013d6ff591e4e	42	ccd8ac5ce51d687a65b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6dcb91f6e987079468e6	42	ccdd93e22ce8fde4aefa	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8b3a31b1df100087704c	42	2ee37f6a5b5afdba3728	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
175bd6ca7da15cd5a63e	42	81f8ac02e568df5d9552	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4fa6e2e476a1f5b1fb15	42	4a894a4948ebc3d652ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
666d1a1bb5bd0d222628	42	3ea70d2daf15198c70a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9e27f0a4cb8667dfd1a6	42	6e3f730a3b6e866694e3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7a1718e5cbcab81393a6	42	cd4dde106f1d53c9aada	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0af9b6812004a8224948	42	011f0cc85f1ad1cc22e6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9cd60d16a185a9b1d470	42	be0802e0a01e96e75eab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
364d762d2b7ba7ea1608	42	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
74d0c3bec1709027cccc	42	042f404c813aed6ee038	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d2ae2a934ef287e75326	42	9538a1857d657c0f004b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a41ad67a14b07cfc7581	42	d74ad31e1cd4b3e1a678	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
16d040d0bca390f9c97e	42	998e79a5a2df8f596fa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9bd3e67ee56139e9d0ab	42	bd6f2ca2ed437f9499d2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f01b787f6ac96e6df902	2007	ccd8ac5ce51d687a65b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8d7936582d1de844a1de	2007	ccdd93e22ce8fde4aefa	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f033937c84e46431b3d9	2007	2ee37f6a5b5afdba3728	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d76d15cbb35910dec146	2007	81f8ac02e568df5d9552	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4407acce01ee2d60b646	2007	4a894a4948ebc3d652ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
70633ae3f8e1d1e708da	2007	3ea70d2daf15198c70a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4517013cce31eb12c37e	2007	1c95850b2ece593eeb02	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a1da6d54cdbb22e7f598	2007	cd4dde106f1d53c9aada	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d9a9793a483ce25ad8f3	2007	011f0cc85f1ad1cc22e6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e5a06a716d7d8b4f33a7	2007	be0802e0a01e96e75eab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
68d52b02c57ffe302d85	2007	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f93b8087c224bc693374	2007	042f404c813aed6ee038	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4bdaaa19fc57254fc1bb	2007	9538a1857d657c0f004b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7f9afbf71116e5e52bc0	2007	d74ad31e1cd4b3e1a678	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f93d1ee171e933fe3cb3	2007	998e79a5a2df8f596fa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a09ac3b2752240b484f3	2007	bd6f2ca2ed437f9499d2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ea1d5a0d2da58300ac45	2038	ccd8ac5ce51d687a65b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
869397ac6fd1feeaf10c	2038	ccdd93e22ce8fde4aefa	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
428500e527c62256a461	2038	2ee37f6a5b5afdba3728	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1edaee7843092e283360	2038	81f8ac02e568df5d9552	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aac2413535bf52cb0556	2038	4a894a4948ebc3d652ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c1c77daae6fad91ddb5f	2038	3ea70d2daf15198c70a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ca76b2cddaa47ecd792d	2038	6e3f730a3b6e866694e3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
688c9e9938d132d1342d	2038	cd4dde106f1d53c9aada	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
de89c50bf247db67f507	2038	011f0cc85f1ad1cc22e6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
74332811ecd7637b70db	2038	be0802e0a01e96e75eab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
45cff309a8cd98074a1e	2038	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d97e0ac03c27651735d6	2038	042f404c813aed6ee038	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c8fa940dd57de469111b	2038	9538a1857d657c0f004b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0be2f50246f0eb979bdd	2038	d74ad31e1cd4b3e1a678	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9a3bda82ed0c38ddd902	2038	998e79a5a2df8f596fa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1f2c3d30d6c61b7a6cfe	2038	bd6f2ca2ed437f9499d2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
40c79449016f29cdd446	864	ccd8ac5ce51d687a65b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
95fd971b041e7f738886	864	ccdd93e22ce8fde4aefa	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e5bf17d10119246b2248	864	2ee37f6a5b5afdba3728	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
67b0d96eed6f7fffbab5	864	81f8ac02e568df5d9552	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2ff71d1115a7f329f98c	864	4a894a4948ebc3d652ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e91e71219c884743cd22	864	3ea70d2daf15198c70a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0b604d34a0c31fdc5716	864	937d339e9ff57e16d3c9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
58c62fc86d035f4a759c	864	cd4dde106f1d53c9aada	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b74519bb589d0fe67555	864	011f0cc85f1ad1cc22e6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3eddc872ff24138bc4ab	864	be0802e0a01e96e75eab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
47ac53e5d4516efb2973	864	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
42758d010ef5d8098dc8	864	042f404c813aed6ee038	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7254d8f55d38c60b2ba3	864	9538a1857d657c0f004b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4c2fd3629a0a02652887	864	d74ad31e1cd4b3e1a678	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7e05f60fc48781b81387	864	998e79a5a2df8f596fa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c93867a31d0ff7532edd	864	bd6f2ca2ed437f9499d2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
43866362173aa501b216	1697	7d66aa886cfe3a72f86b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
365c50e05935344779a4	1697	1993cf9763443fce0f82	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7dfb411df0a9aa70eebc	1697	9a01025dd44aaf5de999	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
33f2159062e0a6146f02	1697	4901021e0e8f2486bbf9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
739003004774d186b950	1697	17be688be816f0468b05	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
35fefe13472203495ca1	1697	b1708239145b8854c8a3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
530edfd42277a285061f	1697	937d339e9ff57e16d3c9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
06f0e756a5017e443a28	1697	7dc8596f01b7d5f27d2f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
190f4a5edbe2e46333a5	1697	58f247c4f27334049744	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2cfd16bf75944b1150cf	1697	417a1d3104a891ccfa38	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dd8114a8772f54f98ee9	1697	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b00b157fa08e0c3e5535	1697	fc618b2ec9d2df99322d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
146f6ba05e0997316004	1697	52a0c7618351d4fe83e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3204efbe881d506ff75b	1697	942711e71cb8fb753a49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
27d2790e4676d46d1ffa	1697	332ec0589e88163cdbc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
74cc27e8ea5b21e07798	1697	c831500edcf62a572d96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
141f0de58978120add21	1082	d8bb7e8a8ebe34b835c3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6e6cb2bb124c05e6a63c	1082	1993cf9763443fce0f82	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
da51b8eddbaef8f9a2c6	1082	9a01025dd44aaf5de999	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b8cfd6b97c3d502c3209	1082	4901021e0e8f2486bbf9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d4eaecaf052fc0e33c2f	1082	17be688be816f0468b05	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
de7b6a70a869e92057b4	1082	b1708239145b8854c8a3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dc71c4cb98606cd30c07	1082	62a6aa3edf4fe2dee2d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a3f0c264ce4b866c27e7	1082	7dc8596f01b7d5f27d2f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6a536823a9f92b9701c7	1082	58f247c4f27334049744	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
17c84843813eb72bf5eb	1082	417a1d3104a891ccfa38	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f27c23b9db680f277f78	1082	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
36a944a4c1358fcecf03	1082	fc618b2ec9d2df99322d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
75d6846961aa142a88b2	1082	52a0c7618351d4fe83e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
796d93af210b145b6554	1082	942711e71cb8fb753a49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aee4d59372f16e4555f2	1082	332ec0589e88163cdbc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
06ede7eb4b0f4714f9e0	1082	c831500edcf62a572d96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8fa227649975b9d53a3c	2186	7d66aa886cfe3a72f86b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c3ef952fa30c6407d299	2186	1993cf9763443fce0f82	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d671f31756c43fec87ef	2186	9a01025dd44aaf5de999	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
113f41d77fd6d713e0f1	2186	4901021e0e8f2486bbf9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8abc208c65237e3147bb	2186	17be688be816f0468b05	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8b48cddd9c99c3ba32ab	2186	b1708239145b8854c8a3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c8e5c93f72c0cb490774	2186	6e3f730a3b6e866694e3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
18555d9b7b5c72e85811	2186	7dc8596f01b7d5f27d2f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a217b4746d6749588e71	2186	58f247c4f27334049744	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8e14eddac75d31cd38b3	2186	417a1d3104a891ccfa38	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bb129b4dc050bfd3aeba	2186	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
943e988fc1b6795ad6c9	2186	fc618b2ec9d2df99322d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c74772ea51cedd15d515	2186	52a0c7618351d4fe83e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
65e23c6cb99ea74fc274	2186	942711e71cb8fb753a49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7087d5ab6ac2be17236d	2186	332ec0589e88163cdbc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8e1ee1b18df3b76a7e3c	2186	c831500edcf62a572d96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5e4d74925adc08f5a318	2362	d8bb7e8a8ebe34b835c3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
63a37bfd0157903e5545	2362	1993cf9763443fce0f82	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7a02855f31571ef0fe8b	2362	9a01025dd44aaf5de999	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9029a88ba902f3afe7ba	2362	4901021e0e8f2486bbf9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
21a2ebfc962a9c766aae	2362	17be688be816f0468b05	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9a477e30d43dd23ef185	2362	b1708239145b8854c8a3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d259d5d215a6cec8c549	2362	62a6aa3edf4fe2dee2d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
771f3897be7d7fa27e40	2362	7dc8596f01b7d5f27d2f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
30e10afcc5d82a654bdf	2362	58f247c4f27334049744	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
34b64f3a44da099ef5b1	2362	417a1d3104a891ccfa38	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e05969ef48ad692eeba3	2362	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2eed002d1fa77db95bf7	2362	fc618b2ec9d2df99322d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1e271e73b5aa81156d9f	2362	52a0c7618351d4fe83e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
43ef86050ad95555e5d7	2362	942711e71cb8fb753a49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b9a89b1477772ae631c6	2362	332ec0589e88163cdbc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b1e74de1bc8c866fe147	2362	c831500edcf62a572d96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
410edeb28ed4d249addd	2569	d8bb7e8a8ebe34b835c3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bb30bffacd5855084704	2569	1993cf9763443fce0f82	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ad3d3d59edd9ef612577	2569	9a01025dd44aaf5de999	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
42200ae9c5777794a696	2569	4901021e0e8f2486bbf9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b8be3119e241c95790a1	2569	17be688be816f0468b05	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9dde4f94ddb3b71cd926	2569	b1708239145b8854c8a3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ea6984b3d29921d9bd19	2569	937d339e9ff57e16d3c9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b913f6ddca8c6a1ef4eb	2569	7dc8596f01b7d5f27d2f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0e9fa56965485973474e	2569	58f247c4f27334049744	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
22b709f0b13d3a386fae	2569	417a1d3104a891ccfa38	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dd523c85fb3f3df73600	2569	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
24e156d9411f0dc47e88	2569	fc618b2ec9d2df99322d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c097b1dd6413847b818a	2569	52a0c7618351d4fe83e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ca5d5d4b09852c3f77c3	2569	942711e71cb8fb753a49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c7403b25e277b88f934f	2569	332ec0589e88163cdbc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aca18633e4b48c65222c	2569	c831500edcf62a572d96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cefdf1b8a5d5c8e86c9f	2236	d8bb7e8a8ebe34b835c3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f058e02ea48633f5f4b0	2236	1993cf9763443fce0f82	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eb3ee2bef9802bf1da24	2236	9a01025dd44aaf5de999	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d042a9eb60789be089c5	2236	4901021e0e8f2486bbf9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8f6745f138152602d4cc	2236	17be688be816f0468b05	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
520e52bb97066f36bb48	2236	b1708239145b8854c8a3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0d4d571dc2db60c84327	2236	62a6aa3edf4fe2dee2d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f614d1ccfc0f56d71fa9	2236	7dc8596f01b7d5f27d2f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
65008f69892574aa7566	2236	58f247c4f27334049744	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7292951d114debb5eebc	2236	417a1d3104a891ccfa38	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c22938ee30c76909bf70	2236	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e408ad8565ad841e53d1	2236	fc618b2ec9d2df99322d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7eb351b8d64a2c28cd59	2236	52a0c7618351d4fe83e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a97f4cc64dd516129d0a	2236	942711e71cb8fb753a49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
86738d55d84fef9833d0	2236	332ec0589e88163cdbc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
75684c7f28246d4ec380	2236	c831500edcf62a572d96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
52006166e41b47fa3418	2040	d8bb7e8a8ebe34b835c3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f92e4b777fc2a8ec226a	2040	1993cf9763443fce0f82	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cd46342f1f44f1e7b792	2040	9a01025dd44aaf5de999	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e5bfaefeb24df1d8da53	2040	4901021e0e8f2486bbf9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6d07997a2c8075fed07f	2040	17be688be816f0468b05	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee4dfbda5e858f2dd438	2040	b1708239145b8854c8a3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8d0c82658a3a5e51df9c	2040	6e3f730a3b6e866694e3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dca1b1d476f1c2a3937f	2040	7dc8596f01b7d5f27d2f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b514859bdc4c1f5544e6	2040	58f247c4f27334049744	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
063ba0acf2cf57b37d29	2040	417a1d3104a891ccfa38	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d30dc9360274ca0a1b96	2040	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7ce7084f0068f6a61181	2040	fc618b2ec9d2df99322d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1fb7acfa7d85194d3c4c	2040	52a0c7618351d4fe83e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7e8be2224cf0233bcdca	2040	942711e71cb8fb753a49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0e607d2109c593d10a23	2040	332ec0589e88163cdbc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d4430e255014f74031aa	2040	c831500edcf62a572d96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8b32909d00f5832c47ae	1250	d8bb7e8a8ebe34b835c3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
76c2a67b5f2156c27a15	1250	1993cf9763443fce0f82	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dca68699502cb3fcc313	1250	9a01025dd44aaf5de999	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1a7b06a373a94eb31c14	1250	4901021e0e8f2486bbf9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9d89b3bcfd49358e8dbf	1250	17be688be816f0468b05	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2726e110e674ab811d47	1250	b1708239145b8854c8a3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bdf32c193d27950c7816	1250	62a6aa3edf4fe2dee2d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3ccf84722b76546fc365	1250	7dc8596f01b7d5f27d2f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5caed3b14c0c3b5b6e1b	1250	58f247c4f27334049744	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
96f713452ac539b63df0	1250	417a1d3104a891ccfa38	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e56f140cc94f23bcef2e	1250	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d5d75e979a4e40a0f1fb	1250	fc618b2ec9d2df99322d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9679cc698fc5d27fe632	1250	52a0c7618351d4fe83e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e08927f3e952ab5ee6ba	1250	942711e71cb8fb753a49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8204d1870cd25819fda9	1250	332ec0589e88163cdbc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b520726a5e4eebc0bbe8	1250	c831500edcf62a572d96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9285b0cfa3b5d999da17	994	d8bb7e8a8ebe34b835c3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
060d86d00d6aa6c4f29c	994	1993cf9763443fce0f82	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bad70857c10c9d866057	994	9a01025dd44aaf5de999	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1c25552ed00477dd4503	994	4901021e0e8f2486bbf9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eef946d1652823e2003b	994	17be688be816f0468b05	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a220ad34f02269c29b5a	994	b1708239145b8854c8a3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
93cdaddaa447babe8cf6	994	937d339e9ff57e16d3c9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
63a0d11f5cf7385749a6	994	7dc8596f01b7d5f27d2f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e18d4dda27fb8ed5b8ee	994	58f247c4f27334049744	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6e2210b1ab30916e9825	994	417a1d3104a891ccfa38	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2af5190b6c4485b02f47	994	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2ddf9062622a1e5c619d	994	fc618b2ec9d2df99322d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
39480b78e357ab90d935	994	52a0c7618351d4fe83e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
75d8c9e26e338942ab76	994	942711e71cb8fb753a49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9ef0d965735f0d7c7b8b	994	332ec0589e88163cdbc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e9030dac233d41db2b7c	994	c831500edcf62a572d96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
814598a9c58c9f3e52ff	2632	d8bb7e8a8ebe34b835c3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ef82e229a9bc2c974674	2632	1993cf9763443fce0f82	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1bdacf43e19664c487a9	2632	9a01025dd44aaf5de999	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0fc911e89ccfacf6033f	2632	4901021e0e8f2486bbf9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dec33a5086e8cec9ad63	2632	17be688be816f0468b05	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1b4ce0793608e61415c9	2632	b1708239145b8854c8a3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
79c04512a342f6476352	2632	62a6aa3edf4fe2dee2d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
40cb16706fe134e76b16	2632	7dc8596f01b7d5f27d2f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c3da03be6ddc402bda10	2632	58f247c4f27334049744	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e1c3abfdd3920b60e1f7	2632	417a1d3104a891ccfa38	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ef256160470e48e2ca4c	2632	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b4dd8518905f1a967b91	2632	fc618b2ec9d2df99322d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0ea79a3bc4f7d8eabbc5	2632	52a0c7618351d4fe83e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
22a711f83cad0760c434	2632	942711e71cb8fb753a49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f972cd5525fdf887fee1	2632	332ec0589e88163cdbc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f51a96b81bf5dc7620eb	2632	c831500edcf62a572d96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4bb96d392a92577f8c5e	2665	d8bb7e8a8ebe34b835c3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1a2d18104ffb47a0d899	2665	1993cf9763443fce0f82	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9df0c9a2db1dce1a3488	2665	9a01025dd44aaf5de999	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e3d42d204d8f67ddda09	2665	4901021e0e8f2486bbf9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f2f97173ce37203d29dd	2665	17be688be816f0468b05	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c3701a7f2c53d45dda55	2665	b1708239145b8854c8a3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b242be4e900d64ecb653	2665	6e3f730a3b6e866694e3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
48f02e018250b2ff9c47	2665	7dc8596f01b7d5f27d2f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a66faebb0bef0cda71ef	2665	58f247c4f27334049744	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ef11b60576e10e85ef78	2665	417a1d3104a891ccfa38	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
049d24bfc2c07a08a204	2665	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fafe8064885327abe6cc	2665	fc618b2ec9d2df99322d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
be73b33aee66fda79e26	2665	52a0c7618351d4fe83e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
36898db10880945a9cd7	2665	942711e71cb8fb753a49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
252fabd077d1da2e3441	2665	332ec0589e88163cdbc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e7a7a84a3233198d1d4d	2665	c831500edcf62a572d96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
56ddea2c3810d91692ac	2615	d8bb7e8a8ebe34b835c3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8bb91e24e11ddf7e63df	2615	1993cf9763443fce0f82	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4825e2566a3f10a1b9d8	2615	9a01025dd44aaf5de999	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2c644b8f040cb366aad8	2615	4901021e0e8f2486bbf9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4cb64d2b7eb4e622afba	2615	17be688be816f0468b05	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
119beb556d515a7da3e0	2615	b1708239145b8854c8a3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
178b1c5c09789940c505	2615	6e3f730a3b6e866694e3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8b38ad9b0e4da98455f4	2615	7dc8596f01b7d5f27d2f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
de41cf8e24d4ab171df9	2615	58f247c4f27334049744	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1eebbaab56725f3bbb12	2615	417a1d3104a891ccfa38	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8db9690c8a40a54dc43b	2615	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b0251a7ca850e67c8909	2615	fc618b2ec9d2df99322d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f818f4cf7945f209bfea	2615	52a0c7618351d4fe83e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
008dccfdb9521241d09a	2615	942711e71cb8fb753a49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
33ad2408a9569c7517ca	2615	332ec0589e88163cdbc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
685981d348f14ef24019	2615	c831500edcf62a572d96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
578cfdb84e1b46d17e28	866	7d66aa886cfe3a72f86b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b443e9e763e44b28ba4c	866	1993cf9763443fce0f82	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7f62e5df0d2fdf83ceaa	866	9a01025dd44aaf5de999	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9471092b774600af3933	866	4901021e0e8f2486bbf9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b17fe9d802f4c4024c6e	866	17be688be816f0468b05	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7fb218128f1faf1ddc3c	866	b1708239145b8854c8a3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f746b0ecc24f262e82bb	866	1c95850b2ece593eeb02	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9fcc33aad6f4f7a11134	866	7dc8596f01b7d5f27d2f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0568deb114265fef7f2f	866	58f247c4f27334049744	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
899b604a945f47f0e82c	866	417a1d3104a891ccfa38	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6b629cc570f73173d6d8	866	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
727dc8113e879dc005bd	866	fc618b2ec9d2df99322d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
83b2183cc7fdec328e7e	866	52a0c7618351d4fe83e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
da5cf2f5b45795d64090	866	942711e71cb8fb753a49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8b54d5a14d6a14cda897	866	332ec0589e88163cdbc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
db6af7b1cf66fad05a59	866	c831500edcf62a572d96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2b4ba3c832e603743df0	2140	d8bb7e8a8ebe34b835c3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
be465376209c21736f46	2140	1993cf9763443fce0f82	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b19cedc9de2bb2c7980f	2140	9a01025dd44aaf5de999	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
49591921f9f837a30671	2140	4901021e0e8f2486bbf9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
30d3ed301afee2ac8279	2140	17be688be816f0468b05	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fe993381e02a23dbb9dd	2140	b1708239145b8854c8a3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f4e3bea5d73ef4df60d9	2140	1c95850b2ece593eeb02	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7d0b29ca433b7c50be16	2140	7dc8596f01b7d5f27d2f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
71a5ccd036b20a431df7	2140	58f247c4f27334049744	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6af4fab6282ea28bffac	2140	417a1d3104a891ccfa38	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f1953b6c90e6147594d6	2140	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
00bbe3cbff5cd16d4af0	2140	fc618b2ec9d2df99322d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c109b450abc038d21bcc	2140	52a0c7618351d4fe83e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
788e3bfe985a69a389e0	2140	942711e71cb8fb753a49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fe9d7efb83fb1edcba6f	2140	332ec0589e88163cdbc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a4a5a476abfe3540e4a8	2140	c831500edcf62a572d96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e6f12121a381073b9c82	2145	7d66aa886cfe3a72f86b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2ad15660a85c864fdcd0	2145	1993cf9763443fce0f82	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d95321c9eb1c0fd61f0a	2145	9a01025dd44aaf5de999	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3dc50e30258cfe39ad1c	2145	4901021e0e8f2486bbf9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
530d1e77b18ebddbea0b	2145	17be688be816f0468b05	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f10c8d9df5ce86d70a96	2145	b1708239145b8854c8a3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9ad35213ed68d635a0a3	2145	6e3f730a3b6e866694e3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
106186459cd16ce6d6f8	2145	7dc8596f01b7d5f27d2f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7bd7c1e185db432d1ade	2145	58f247c4f27334049744	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3bf6f70f769c4a1373e1	2145	417a1d3104a891ccfa38	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
451b8132626bd4315b27	2145	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
13947b15494b7b147945	2145	fc618b2ec9d2df99322d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c53773e9a34e44808966	2145	52a0c7618351d4fe83e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
80bd7c81c8273c96a59e	2145	942711e71cb8fb753a49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d27ef8bcc26081b69de2	2145	332ec0589e88163cdbc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
014e1894239013949898	2145	c831500edcf62a572d96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c90be12faa39c0cbc9c4	1904	7d66aa886cfe3a72f86b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c418de44ce032cbc2a9d	1904	1993cf9763443fce0f82	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
635d71f75cd2b8068a39	1904	9a01025dd44aaf5de999	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2e3d7751f2acee94d8cd	1904	4901021e0e8f2486bbf9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dcc5121d6a1fbd963b32	1904	17be688be816f0468b05	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6ec09b8f8e6da52fd8c1	1904	b1708239145b8854c8a3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee361e8874a41f6baa63	1904	6e3f730a3b6e866694e3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c4088f553c13e17b18e7	1904	7dc8596f01b7d5f27d2f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
50e6b2cfe5798bac8b58	1904	58f247c4f27334049744	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1ac11d7808d093470035	1904	417a1d3104a891ccfa38	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8ed2d85f67b8f2ae5675	1904	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c26d5716e20406d0c335	1904	fc618b2ec9d2df99322d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eb4940226c36bd2c32df	1904	52a0c7618351d4fe83e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b0a52faa3f6280764eb1	1904	942711e71cb8fb753a49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c292cdc19d7ec1d7d518	1904	332ec0589e88163cdbc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
177f8c7392ad051daa45	1904	c831500edcf62a572d96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d8b8ffbf8f24b5d06296	973	d8bb7e8a8ebe34b835c3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
28374aa5e925a381d103	973	1993cf9763443fce0f82	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1360c8e878ddbd8374f2	973	9a01025dd44aaf5de999	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
14488d60a832ee4a28d5	973	4901021e0e8f2486bbf9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1633077bb27c8f20493a	973	17be688be816f0468b05	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bd1e1520e35814a1b889	973	b1708239145b8854c8a3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d479f823aa46fa671264	973	6e3f730a3b6e866694e3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
83e67fcbcba75799c248	973	7dc8596f01b7d5f27d2f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7b0f793e3ad9e92bbbef	973	58f247c4f27334049744	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1cc2f951086433af0269	973	417a1d3104a891ccfa38	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
97abbd29ddbec6ddafa4	973	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6b2eef1666b00ca9d557	973	fc618b2ec9d2df99322d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d320a176f68de6b0c27d	973	52a0c7618351d4fe83e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a014096a8d46dafb1227	973	942711e71cb8fb753a49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
32de831cf67ad970250b	973	332ec0589e88163cdbc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f7389d005d1e6103519c	973	c831500edcf62a572d96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cfbaf1f80dfa4a86063d	1140	d8bb7e8a8ebe34b835c3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5976d45df05719eb8c57	1140	1993cf9763443fce0f82	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d383589804906e548d22	1140	9a01025dd44aaf5de999	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
47ae077d12dff2d406b4	1140	4901021e0e8f2486bbf9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
995131c47fef161c677e	1140	17be688be816f0468b05	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b443cde5df6ace2e0031	1140	b1708239145b8854c8a3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6939eb348bdc30019980	1140	6e3f730a3b6e866694e3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f390066424a4ab82fbc8	1140	7dc8596f01b7d5f27d2f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
256d47de3e61abcb6d85	1140	58f247c4f27334049744	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1b2196f5e27c3fa7ec90	1140	417a1d3104a891ccfa38	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8b0c206d7ea5ce3e5066	1140	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
97272f8b000dbb7594c0	1140	fc618b2ec9d2df99322d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fc62b58e3ec28432a52a	1140	52a0c7618351d4fe83e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
856d0838d6a41a0fec2e	1140	942711e71cb8fb753a49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
051b73b2de4cc1c10068	1140	332ec0589e88163cdbc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bda9f8088db94005de5c	1140	c831500edcf62a572d96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b2002af122496707ac55	912	d8bb7e8a8ebe34b835c3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
73df3ef8b2357873a154	912	1993cf9763443fce0f82	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
649a3640e90b497ebbdd	912	9a01025dd44aaf5de999	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
62ec45f8f27dd2103945	912	4901021e0e8f2486bbf9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6653bfba080c44e55339	912	17be688be816f0468b05	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
56edfefebf31c3522975	912	b1708239145b8854c8a3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
970b0e7290c51804f071	912	62a6aa3edf4fe2dee2d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fb154e6515fa0008fc02	912	7dc8596f01b7d5f27d2f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2eb83234b8d5f592a2fb	912	58f247c4f27334049744	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
38da8b1919024982f8d5	912	417a1d3104a891ccfa38	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
36f08d2308b0f42231a1	912	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
810ef53cfc030572787c	912	fc618b2ec9d2df99322d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0514869b564f4cb01021	912	52a0c7618351d4fe83e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
77f5854cd10080025492	912	942711e71cb8fb753a49	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7fe6d64168e13033af85	912	332ec0589e88163cdbc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ece3189c71cafccff1c9	912	c831500edcf62a572d96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
92132e74153138f0236e	2204	0f9cf77856a913b30c8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b1351b92e0db8477bc86	2204	67a7363447e757702674	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
78ada46c032c40c35a04	2204	524cb475f7eec63c3736	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
87c9e68690f2c6690bc1	2204	a09b6ca1c047b268f7cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fb72600b880536e0bee7	2204	6c05b760b3d8e03e72a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cec54bae8b42250786e6	2204	ffecd66e5100127c4042	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fe94674593f4b4715589	2204	6e3f730a3b6e866694e3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d27d2511d7f6b3bd3bf1	2204	dac65aeaaed135b88dba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
95926016ff42d97f2d9b	2204	4883160b03bc92b91640	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
be5436b6daf6dce13bc8	2204	d87f1150ff27c1efbcc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bf4e08de27de8b5eeddf	2204	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8b7fd62530b5abd35be7	2204	9b8fac85cf8101939f27	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c8d247dc5009c0842605	2204	0716aa5bb459f01a8e41	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f7cca0277e019ffe55f1	2204	1bfba56a05f5800f2691	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e124844fb2b6026e3f79	2204	fe2dc648d4dc41df5d30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
693c0b15d7dfb25f9d68	2204	42c2357c9e79a8d2bd9e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
58383fc4f8fb6c3a7fd9	2586	140cb8e7fd7f918244a8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7ccdbcdc9c43ebce5fa4	2586	67a7363447e757702674	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ade76bc8fee654c5f69c	2586	524cb475f7eec63c3736	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
39ef8ec8183c8fa39cbd	2586	a09b6ca1c047b268f7cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
66cd7b7fe313be825be0	2586	6c05b760b3d8e03e72a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5ecde824430f328aa3e9	2586	ffecd66e5100127c4042	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3b98db00b536de2003d2	2586	1c95850b2ece593eeb02	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cb02d10ab3d47b06ec17	2586	dac65aeaaed135b88dba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
267db048b8485602a7ab	2586	4883160b03bc92b91640	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
64ef4d1a5b872d4dd592	2586	d87f1150ff27c1efbcc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d7032a312acccda4cfeb	2586	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
93b096c522a3c496e940	2586	9b8fac85cf8101939f27	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f7137ec3efa885b7b391	2586	0716aa5bb459f01a8e41	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4dae6bc9801abf72ce03	2586	1bfba56a05f5800f2691	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0a49feb99826655520ce	2586	fe2dc648d4dc41df5d30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3a7a728d905b5f2c7201	2586	42c2357c9e79a8d2bd9e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
94087aeeede676b6c2ad	15	140cb8e7fd7f918244a8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
52d6176913157baa400a	15	67a7363447e757702674	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
23f46012183fa77dbc8f	15	524cb475f7eec63c3736	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a4985c49f3ef24d124b2	15	a09b6ca1c047b268f7cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fe2ab86bcbb4854283f6	15	6c05b760b3d8e03e72a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7afa707217a562cfc3d1	15	ffecd66e5100127c4042	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f95f7bb84edfc7b263ee	15	1c95850b2ece593eeb02	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3fad6d25df58bf9f0b43	15	dac65aeaaed135b88dba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
48d8fa8822365dcff88d	15	4883160b03bc92b91640	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
692badac2634753a3dbc	15	d87f1150ff27c1efbcc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
469cea7da65b9d19db23	15	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
20347fe16ee2e7d2a603	15	9b8fac85cf8101939f27	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5e1e5e0c99be5490da75	15	0716aa5bb459f01a8e41	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
463afb3507134a701cc6	15	1bfba56a05f5800f2691	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ec26571d134cae7a1364	15	fe2dc648d4dc41df5d30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee102f6ac2f54eb6d73c	15	42c2357c9e79a8d2bd9e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1e205027b4059da975c1	22	140cb8e7fd7f918244a8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
84ff1a15884f77bc7ebb	22	67a7363447e757702674	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d9cfe6f3a8f299dffa1b	22	524cb475f7eec63c3736	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eb079ebe6760747c2d1d	22	a09b6ca1c047b268f7cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
56c371e65e440ba82a42	22	6c05b760b3d8e03e72a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
23dce780e58da80d7923	22	ffecd66e5100127c4042	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cd7c86185dfd5679fe42	22	937d339e9ff57e16d3c9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2c6ec1c92279f67668f0	22	dac65aeaaed135b88dba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fbd39cebe535dad76edf	22	4883160b03bc92b91640	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
02beb885aa0b715f0a30	22	d87f1150ff27c1efbcc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8d38b0693821f953315f	22	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aae5b27e8ff9c1834841	22	9b8fac85cf8101939f27	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
16edfe813f87c7024bd6	22	0716aa5bb459f01a8e41	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8ea5113dc7d4737d4bd4	22	1bfba56a05f5800f2691	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cfc55a8a44de595a7ea8	22	fe2dc648d4dc41df5d30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b852ccd92143f6ffe71c	22	42c2357c9e79a8d2bd9e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ddd4bf0e2da039a2aabc	2570	0f9cf77856a913b30c8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7377be763cea242e8c9f	2570	67a7363447e757702674	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f6609d2b3cbb8b699e1a	2570	524cb475f7eec63c3736	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b30dc5431a47e7a4d3d5	2570	a09b6ca1c047b268f7cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2cb56f1f01c8d66f94a9	2570	6c05b760b3d8e03e72a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a1451a42db4a201c31cf	2570	ffecd66e5100127c4042	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
df5d3f7b7bbf741a8c59	2570	1c95850b2ece593eeb02	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c51c6ec99ab41a5e8537	2570	dac65aeaaed135b88dba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8e43028caf1b2128364a	2570	4883160b03bc92b91640	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
36fb293dded61afca1bb	2570	d87f1150ff27c1efbcc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
79e3d1f7a3e8cdbc9dd6	2570	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b5f79ba826a7eb870d9c	2570	9b8fac85cf8101939f27	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cba5edb7b35efa9bbcc7	2570	0716aa5bb459f01a8e41	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
daa958510035b2af83ec	2570	1bfba56a05f5800f2691	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9ebc84072c0e8f81ecc6	2570	fe2dc648d4dc41df5d30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fc028a4590bd1f71b2c5	2570	42c2357c9e79a8d2bd9e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6324e2ac8586d9b11a93	2315	0f9cf77856a913b30c8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
daa44df7287893baabfc	2315	67a7363447e757702674	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1227d93ef3a2e487318a	2315	524cb475f7eec63c3736	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
81378b38af765fc0a40b	2315	a09b6ca1c047b268f7cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2e0456c1da7b106b11ad	2315	6c05b760b3d8e03e72a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1990dcbe278aa551eb95	2315	ffecd66e5100127c4042	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8b28064f323da4837e27	2315	937d339e9ff57e16d3c9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5e9552a0ac4172dc7ca4	2315	dac65aeaaed135b88dba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9077c81594d5ec32075f	2315	4883160b03bc92b91640	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bd987be04a6f9a117317	2315	d87f1150ff27c1efbcc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8833f27b7254bd8ed977	2315	67cb84d4c26e5522b07c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
74f8d75e13d9c634c53d	2315	9b8fac85cf8101939f27	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0c7c045479cc33290fa1	2315	0716aa5bb459f01a8e41	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
50a3fae2b836de5da448	2315	1bfba56a05f5800f2691	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7f411d785feba09a250d	2315	fe2dc648d4dc41df5d30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f15cd51f3ddaf9b3cbc5	2315	42c2357c9e79a8d2bd9e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ed4abc15ea68e849f90f	34	140cb8e7fd7f918244a8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aff515889d52b37efd40	34	67a7363447e757702674	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
61d42f366480b83f6423	34	524cb475f7eec63c3736	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
73e4a9c19e6181775fb8	34	a09b6ca1c047b268f7cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e4f273a96373367cc6e4	34	6c05b760b3d8e03e72a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
64753a1f3e6e553c6617	34	ffecd66e5100127c4042	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3d5cf59e28d298f59648	34	1c95850b2ece593eeb02	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
85c6a8f4b3634f4b833d	34	dac65aeaaed135b88dba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dbd283c7d915a8bec7de	34	4883160b03bc92b91640	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
12930cdcb88fbce81fed	34	d87f1150ff27c1efbcc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7b88ca6408ed41580120	34	67cb84d4c26e5522b07c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a8a3ae340e5c9a8cb2c1	34	9b8fac85cf8101939f27	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0b99dfa1aac2bc340370	34	0716aa5bb459f01a8e41	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d92852c0d20aca0f5e82	34	1bfba56a05f5800f2691	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e67118ede47dd4ed3cb8	34	fe2dc648d4dc41df5d30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
681d38cc934353b3ffe6	34	42c2357c9e79a8d2bd9e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5ead6d754bf712ca9ac1	2139	0f9cf77856a913b30c8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee00f1895c576c7affa6	2139	67a7363447e757702674	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d3e07541598173cb6f54	2139	524cb475f7eec63c3736	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
39ce957e87d2ecf401f8	2139	a09b6ca1c047b268f7cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
024c3eb918fbaa075296	2139	6c05b760b3d8e03e72a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
45bd0818ad8d3f23358f	2139	ffecd66e5100127c4042	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2b5e1e956c7ac810bea3	2139	6e3f730a3b6e866694e3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e5cbcbf4610520984a1c	2139	dac65aeaaed135b88dba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
96b467ba93b3752996a3	2139	4883160b03bc92b91640	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
007a6faa109dc5a4711d	2139	d87f1150ff27c1efbcc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
99f9e672c726771dc9b7	2139	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c6c68b107d8c4637d9f2	2139	9b8fac85cf8101939f27	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8a68e5ae202c494d9f9e	2139	0716aa5bb459f01a8e41	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c3be7a48391c1dba3c92	2139	1bfba56a05f5800f2691	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a956820797f8040c830e	2139	fe2dc648d4dc41df5d30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
801edc9f81f3b557aa46	2139	42c2357c9e79a8d2bd9e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
869f233b8444c11ddf4b	2374	0f9cf77856a913b30c8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7b3d1ebe3a771ffc4f94	2374	67a7363447e757702674	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e643509deb3953254c6d	2374	524cb475f7eec63c3736	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bc0cf2e0c3f90bff888b	2374	a09b6ca1c047b268f7cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
99172187705c69e14219	2374	6c05b760b3d8e03e72a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
16e7a6869d3f1a2a0114	2374	ffecd66e5100127c4042	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d78eceb4ef9c0d1063e8	2374	937d339e9ff57e16d3c9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
933d06aa7823be7e7451	2374	dac65aeaaed135b88dba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f0823f36f299369adfb8	2374	4883160b03bc92b91640	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6cbcd9dea989d3ecd366	2374	d87f1150ff27c1efbcc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
953b9a3a097b49719466	2374	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6dacfb4da55e10ca2c27	2374	9b8fac85cf8101939f27	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
41ad2fd3254c970fe082	2374	0716aa5bb459f01a8e41	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
02eca5b308eef8801276	2374	1bfba56a05f5800f2691	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
76693ad24986bcb570c9	2374	fe2dc648d4dc41df5d30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c24e63488bdf442724c2	2374	42c2357c9e79a8d2bd9e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6f80007429dd85136b39	2364	0f9cf77856a913b30c8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8d56662857c63c64c564	2364	67a7363447e757702674	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
71f83fb466cc819f68b6	2364	524cb475f7eec63c3736	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e69a7219fb8693f0cb4c	2364	a09b6ca1c047b268f7cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cc01dc317fdec668e861	2364	6c05b760b3d8e03e72a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e9525bed2a555fbce529	2364	ffecd66e5100127c4042	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ead9eae111b357ec0e30	2364	1c95850b2ece593eeb02	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8889ca78cdde4a87dcfe	2364	dac65aeaaed135b88dba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5ebf8658cf941932984d	2364	4883160b03bc92b91640	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
057ac82af364264a1398	2364	d87f1150ff27c1efbcc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee101442f9ab86d53f0b	2364	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
54015de087f174843f01	2364	9b8fac85cf8101939f27	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
27027c6f8c3d12c9cf7e	2364	0716aa5bb459f01a8e41	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
da57dd8e5ff5804d68dd	2364	1bfba56a05f5800f2691	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f9e68b329c4602d58f3c	2364	fe2dc648d4dc41df5d30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7e33fcf397fc3bcdb791	2364	42c2357c9e79a8d2bd9e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b7ba266e2fdf8df0dd5d	2658	0f9cf77856a913b30c8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a01e8ab56240092bc485	2658	67a7363447e757702674	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a6fad6b12af04d6a9338	2658	524cb475f7eec63c3736	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
827326a16b78dc6559de	2658	a09b6ca1c047b268f7cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ba9d7a5ff3f81cd5f25a	2658	6c05b760b3d8e03e72a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
693bc4764d3f1b204caa	2658	ffecd66e5100127c4042	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a7b7807cb3920c40249b	2658	6e3f730a3b6e866694e3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
350546e9fceee84d4a3d	2658	dac65aeaaed135b88dba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d4055ec770138ed0a165	2658	4883160b03bc92b91640	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1af47cd9f6f5cec8d2b6	2658	d87f1150ff27c1efbcc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
552a58fa74030495348c	2658	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
90211858db9970de13aa	2658	9b8fac85cf8101939f27	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7f222ae274fdaac3ae6a	2658	0716aa5bb459f01a8e41	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8a0dfa8ecc566dfe61d1	2658	1bfba56a05f5800f2691	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e010a57c5bdd4ff7ecba	2658	fe2dc648d4dc41df5d30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
899b759520bff9218f24	2658	42c2357c9e79a8d2bd9e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eb7f5f5aa2003840bdd8	1259	140cb8e7fd7f918244a8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
59cf5f1ee9d3a081dd92	1259	67a7363447e757702674	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b5efdf07464a2b0f5431	1259	524cb475f7eec63c3736	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a2f3f60ff48aed5b0555	1259	a09b6ca1c047b268f7cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e7c04af5fa37d8365b91	1259	6c05b760b3d8e03e72a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7965c262e309429c6399	1259	ffecd66e5100127c4042	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5f11ae11d2a3e359c984	1259	62a6aa3edf4fe2dee2d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d8a8d1862d68279886e9	1259	dac65aeaaed135b88dba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9a404f45d7e85b83c737	1259	4883160b03bc92b91640	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9353a48e3decc2a9497a	1259	d87f1150ff27c1efbcc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
487f1f509947ccdad33d	1259	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a5a293fd5a92a4498a75	1259	9b8fac85cf8101939f27	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
adec3a8f06646dd2a9e4	1259	0716aa5bb459f01a8e41	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
73b00e0cc093c29caeaf	1259	1bfba56a05f5800f2691	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
45a9e0adef891f6623fa	1259	fe2dc648d4dc41df5d30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7922678d23e1ee920aa2	1259	42c2357c9e79a8d2bd9e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
872d9e124c1d4047408b	2588	0f9cf77856a913b30c8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
66fe3364dd6b0aa44065	2588	67a7363447e757702674	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
747d878d0271d51d8688	2588	524cb475f7eec63c3736	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d64bbfe825b8025bd7a4	2588	a09b6ca1c047b268f7cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
debe03c6128a65569348	2588	6c05b760b3d8e03e72a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d1f49c010f44219470b4	2588	ffecd66e5100127c4042	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
11129bce2622dcf41daa	2588	62a6aa3edf4fe2dee2d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b6ebc34024871e22b909	2588	dac65aeaaed135b88dba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
81b3928eef6744f7c550	2588	4883160b03bc92b91640	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9083cd0f6d3246b457df	2588	d87f1150ff27c1efbcc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e6439f442361e1befdce	2588	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e4cfd099d88019cb91dd	2588	9b8fac85cf8101939f27	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4126dcfa5f8b70c580a3	2588	0716aa5bb459f01a8e41	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
220f2b3afcc68e68452c	2588	1bfba56a05f5800f2691	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1d209c0b23f1ad11e4d1	2588	fe2dc648d4dc41df5d30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
52aaca390e9aa7fbf9b0	2588	42c2357c9e79a8d2bd9e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6fe96a0121102603962b	2344	0f9cf77856a913b30c8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1a22fdb4a1af2e800ee6	2344	67a7363447e757702674	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
35ea9fba8bfe32c631ff	2344	524cb475f7eec63c3736	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ffb71ecef326c9985236	2344	a09b6ca1c047b268f7cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
32a28e7eec6be60f6967	2344	6c05b760b3d8e03e72a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2fbddd34729a69eafc97	2344	ffecd66e5100127c4042	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
af4c376f8a1f4c597ec9	2344	6e3f730a3b6e866694e3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a5c25f82d9a4bfa6642b	2344	dac65aeaaed135b88dba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
99aaa287a8ed057036fc	2344	4883160b03bc92b91640	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
79a25c9f6920a13cd18d	2344	d87f1150ff27c1efbcc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a1a4fbe52e59862e1661	2344	67cb84d4c26e5522b07c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
edf2f55066af1e6a573a	2344	9b8fac85cf8101939f27	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9ad08bf936f2f051726b	2344	0716aa5bb459f01a8e41	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9617745ae13d60a2520a	2344	1bfba56a05f5800f2691	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dcfcad3a66735d83d0c0	2344	fe2dc648d4dc41df5d30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
76233f173079d643f7d3	2344	42c2357c9e79a8d2bd9e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5478e33d1da3119f68d1	41	0f9cf77856a913b30c8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
88c9babe35a6278267a5	41	67a7363447e757702674	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c8ca9cc9c077d35dd72c	41	524cb475f7eec63c3736	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
390a5e6e807aec917dd5	41	a09b6ca1c047b268f7cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6bfa001aa2213cdea9e3	41	6c05b760b3d8e03e72a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
594f4ce91a6b368f1933	41	ffecd66e5100127c4042	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bcab4637347d4c07bfeb	41	937d339e9ff57e16d3c9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a786f7830ac03b0b9baf	41	dac65aeaaed135b88dba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aa0b9691ebe0ca81cf1c	41	4883160b03bc92b91640	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c3c7b4fcc653edec84c9	41	d87f1150ff27c1efbcc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4461f3339b0c3712135e	41	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1eed00b7552635a2fe16	41	9b8fac85cf8101939f27	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9cfe4eb14ac8309435d4	41	0716aa5bb459f01a8e41	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5d1c2258a985e185cbd7	41	1bfba56a05f5800f2691	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
87790e12191f899b22f1	41	fe2dc648d4dc41df5d30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dbfabaf6c989bdf43cc6	41	42c2357c9e79a8d2bd9e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
90fa2c0320a61b477156	2367	0f9cf77856a913b30c8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aec524c0ab91cfcfdd5f	2367	67a7363447e757702674	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0c67f04c744122722e58	2367	524cb475f7eec63c3736	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c1a7036446ac511688a3	2367	a09b6ca1c047b268f7cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8883b0c1462ba7c33aad	2367	6c05b760b3d8e03e72a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dff91b347d2e045f71e1	2367	ffecd66e5100127c4042	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
652898bb239449070bfe	2367	6e3f730a3b6e866694e3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3e31574e05472890a067	2367	dac65aeaaed135b88dba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7b56bc32c736d9267384	2367	4883160b03bc92b91640	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cbc5c1a4762fe58230b8	2367	d87f1150ff27c1efbcc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e3b767af2c1762860321	2367	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
43ff027e7638f675fa7e	2367	9b8fac85cf8101939f27	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6b4850e051cd46e9b422	2367	0716aa5bb459f01a8e41	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fff48802ed04d29a4e36	2367	1bfba56a05f5800f2691	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8a5bcf4d488d8a9b45eb	2367	fe2dc648d4dc41df5d30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
66aa91cae553d8f011ee	2367	42c2357c9e79a8d2bd9e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
df41de139b86e2e40e0c	2399	0f9cf77856a913b30c8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e31fd8a0405151753066	2399	67a7363447e757702674	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5e491f144e3d0f584b25	2399	524cb475f7eec63c3736	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
28875b4303de1e227e1f	2399	a09b6ca1c047b268f7cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0a17281eb90a7b00b5a8	2399	6c05b760b3d8e03e72a7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
930a3d407643ab095747	2399	ffecd66e5100127c4042	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d6eae598688301c66f16	2399	6e3f730a3b6e866694e3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6e655c0872726af3cc93	2399	dac65aeaaed135b88dba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6dacbfa9670f98f8c64c	2399	4883160b03bc92b91640	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
37e4d681dc19ad53f44e	2399	d87f1150ff27c1efbcc2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2870661486727ce33889	2399	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f3b7a4bd7dea31357da8	2399	9b8fac85cf8101939f27	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
26f37d70f0c79fd3c289	2399	0716aa5bb459f01a8e41	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d4b3516a6605a56b59f1	2399	1bfba56a05f5800f2691	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e1582792b2d76b1c4925	2399	fe2dc648d4dc41df5d30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
86cda0d0212b402d0dab	2399	42c2357c9e79a8d2bd9e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7403f4e51b3ddba6d67c	12	c38a7595c8fe93bb3a4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
63a22692d963a6e8acaa	12	922a219336536e10a18d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b97fa35b71908432d847	12	7a6c12bdaf5c6da7b21d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4d4907201d1455b93f99	12	5e3639a7c71b3a997aa8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
727398ca9761e954eccc	12	ad93cb9499239edb202a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6c0d8e32e2de1f0bf52f	12	53fd09a399c9afcd820e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c0408d589c408bdbbb54	12	6e3f730a3b6e866694e3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5742fa0f0f23219dcc15	12	0da1a8e4012fbe1989cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
19b2a0247a8f86074f00	12	83db990ac81dfd8b35c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
31d4c8b6582d108a16d5	12	74b1f4902d76a7737fad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8d1bb83ce918521673c0	12	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c6227233abccb4de749f	12	6128f2c7422b743e9a52	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
394c6316eb829595aa69	12	26f71e77ebb5d7b35f70	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
126120e1573897e4f067	12	3810d251d2a035b1de85	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
126a98287955b46da89f	12	8899967fd022bd010f96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b2de4c854a1b3a1f6bbe	12	f11f70d30bce575096ec	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5079270a62c7c84ff047	861	8b094bf3bca9e805dbab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e24c5daadfb3ffc0d2ff	861	922a219336536e10a18d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
40ba392f34b67b9fa3f7	861	7a6c12bdaf5c6da7b21d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a859233a0c6adad6c185	861	5e3639a7c71b3a997aa8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d34da74d3eb67d040271	861	ad93cb9499239edb202a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2ec5e20937ded80bc374	861	53fd09a399c9afcd820e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dc7cdd07d558dc6cb429	861	1c95850b2ece593eeb02	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
20f7004494ba59edc1a0	861	0da1a8e4012fbe1989cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9faf23e0a2cd1cc09aa0	861	83db990ac81dfd8b35c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
89c3fa5bcfcf917bd935	861	74b1f4902d76a7737fad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b7a9e28fe10ad233da26	861	67cb84d4c26e5522b07c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ad090eab3827c441ae2d	861	6128f2c7422b743e9a52	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9c30a5092f6ced979d7d	861	26f71e77ebb5d7b35f70	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
56d3c2b2836234466b32	861	3810d251d2a035b1de85	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
643f3cfc5330925e6c82	861	8899967fd022bd010f96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4fb44540cef9c4511940	861	f11f70d30bce575096ec	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
966f9287f5d27756d888	1457	8b094bf3bca9e805dbab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3b5563a2d82f58df6853	1457	922a219336536e10a18d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dd299d205e90c3118db0	1457	7a6c12bdaf5c6da7b21d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
17ec64118d03960bbb46	1457	5e3639a7c71b3a997aa8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
768512c9a8c7e68aeca5	1457	ad93cb9499239edb202a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2af14d5a9e593ee5e95a	1457	53fd09a399c9afcd820e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0d5d484cb2512f5f763b	1457	937d339e9ff57e16d3c9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f1578306bcb597e70be8	1457	0da1a8e4012fbe1989cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
282c28729a759f928ba0	1457	83db990ac81dfd8b35c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9c7728abf90ec1705afc	1457	74b1f4902d76a7737fad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
93841817cf68df476069	1457	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2ece5edec7916f35794b	1457	6128f2c7422b743e9a52	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a27832dfd142dc8435a2	1457	26f71e77ebb5d7b35f70	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
59da09659f008d502226	1457	3810d251d2a035b1de85	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e70de50b2fc4d727402e	1457	8899967fd022bd010f96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
202676c2c7fd57337938	1457	f11f70d30bce575096ec	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e5d5916865e9df0f3da3	1599	8b094bf3bca9e805dbab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ac31b4c118206a00ee44	1599	922a219336536e10a18d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fbaa7cba80498e6a0778	1599	7a6c12bdaf5c6da7b21d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1aa24eba1358b4ec41d8	1599	5e3639a7c71b3a997aa8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8dd8766794fa545530d3	1599	ad93cb9499239edb202a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6b6512dcb04dcb86dc72	1599	53fd09a399c9afcd820e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
762efa1ddfcea611fff4	1599	937d339e9ff57e16d3c9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d0846010a4d35806b978	1599	0da1a8e4012fbe1989cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bfe4e82f82122ffeff64	1599	83db990ac81dfd8b35c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
26423321a2146d1b611e	1599	74b1f4902d76a7737fad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d36fce70602ca058f863	1599	67cb84d4c26e5522b07c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1ae48ca8c4f5194372b7	1599	6128f2c7422b743e9a52	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7c6e2ef80fdca2736fc1	1599	26f71e77ebb5d7b35f70	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6f542ac44cca7f121522	1599	3810d251d2a035b1de85	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f59b2ef0f4aaedf06625	1599	8899967fd022bd010f96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0cf5d336d8020c72f730	1599	f11f70d30bce575096ec	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
feab1d35d85492cef357	1769	8b094bf3bca9e805dbab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8535789596aed8f5832f	1769	922a219336536e10a18d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a7aa712807d8869d323a	1769	7a6c12bdaf5c6da7b21d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
428a6d78f4d5dc271e73	1769	5e3639a7c71b3a997aa8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2e74ad8badb26c922d9d	1769	ad93cb9499239edb202a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
791df7f235ead5896c3c	1769	53fd09a399c9afcd820e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
91fecc3081a0f4ef1f79	1769	1c95850b2ece593eeb02	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e025779ce9b02f56526c	1769	0da1a8e4012fbe1989cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
36a3e9a56740f736da59	1769	83db990ac81dfd8b35c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
38a57a67169d7a55503c	1769	74b1f4902d76a7737fad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a1ea183bad7679d99091	1769	67cb84d4c26e5522b07c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
557ffbb0bf361cf9b6e8	1769	6128f2c7422b743e9a52	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1c2d714bba98ee9b6d1f	1769	26f71e77ebb5d7b35f70	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f3fc851011767e79c2e3	1769	3810d251d2a035b1de85	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2d3dd2b761463ef9ce89	1769	8899967fd022bd010f96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fabd516735ce6ae5329c	1769	f11f70d30bce575096ec	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
85709617a8bdc8914ab2	2309	8b094bf3bca9e805dbab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
26591759789c13d9539b	2309	922a219336536e10a18d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7747361f9ac2093c110d	2309	7a6c12bdaf5c6da7b21d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b0e41116ca1770cdf211	2309	5e3639a7c71b3a997aa8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a430c83749fba048ced1	2309	ad93cb9499239edb202a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f5cc4aead215a50dd2b4	2309	53fd09a399c9afcd820e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
214af92edc85797bd2ea	2309	62a6aa3edf4fe2dee2d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3f0d282cf79f3c4599ea	2309	0da1a8e4012fbe1989cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
54d27a6a6066af926013	2309	83db990ac81dfd8b35c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c616d4f9e84cc65aea4c	2309	74b1f4902d76a7737fad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a865b303c98028199c28	2309	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e09c8edf4fc9d34f5853	2309	6128f2c7422b743e9a52	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
03a2366e0a5bee8fcde3	2309	26f71e77ebb5d7b35f70	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3e75cf4bc3f83396b6a4	2309	3810d251d2a035b1de85	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ecd9a9e0708b1f14979d	2309	8899967fd022bd010f96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6894b37a65a43b4ec5a1	2309	f11f70d30bce575096ec	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
932699e476ad9d25ace2	2571	c38a7595c8fe93bb3a4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3ae53ca0ec21e7d49824	2571	922a219336536e10a18d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c793258e13c708ca5a10	2571	7a6c12bdaf5c6da7b21d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
da1798f696ce6d3eb0ef	2571	5e3639a7c71b3a997aa8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
15c3b243d5eac60fde01	2571	ad93cb9499239edb202a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6fbe3ef65f6ba951672d	2571	53fd09a399c9afcd820e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
62a25850a01dcc548111	2571	62a6aa3edf4fe2dee2d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8d7ec74ba380d9d3a9b2	2571	0da1a8e4012fbe1989cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5a3c253acdf654141895	2571	83db990ac81dfd8b35c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
555bcd3e651bb293ea61	2571	74b1f4902d76a7737fad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
42bc36b61399857dbd9c	2571	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
18d49f534f2e28a1731a	2571	6128f2c7422b743e9a52	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a2db75b6833ae7d44d57	2571	26f71e77ebb5d7b35f70	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6b345f463a595002e7e3	2571	3810d251d2a035b1de85	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b473144b3b9840063d4a	2571	8899967fd022bd010f96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f3b85a0edd39429d82b7	2571	f11f70d30bce575096ec	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8b5866f25306e8a49be8	2316	8b094bf3bca9e805dbab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f3d894fda83312bb9430	2316	922a219336536e10a18d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
708f0f0842d833d32ce5	2316	7a6c12bdaf5c6da7b21d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
88dc6b776a53d00a3a15	2316	5e3639a7c71b3a997aa8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0f2a6012358c9ff9f6dd	2316	ad93cb9499239edb202a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f355cffa19e59e8820dc	2316	53fd09a399c9afcd820e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
55d7c5d36e33e9f5dfe0	2316	937d339e9ff57e16d3c9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
06836eacf2ed814f656b	2316	0da1a8e4012fbe1989cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c722d4620b789efe7c01	2316	83db990ac81dfd8b35c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c49c0ce4547582eba75d	2316	74b1f4902d76a7737fad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
002c31d470a165e5448d	2316	67cb84d4c26e5522b07c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a52a650d2d0c1f361c06	2316	6128f2c7422b743e9a52	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6f71ce4cb8490e7aa0b1	2316	26f71e77ebb5d7b35f70	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cf117908d2493210c49e	2316	3810d251d2a035b1de85	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1ee3d3956473ef0b6494	2316	8899967fd022bd010f96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8b2aeaba029497dd5a31	2316	f11f70d30bce575096ec	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9782b325ef669d45999e	2654	8b094bf3bca9e805dbab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
052d569c6692945f867a	2654	922a219336536e10a18d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f3bf015d4cfe627a4b8d	2654	7a6c12bdaf5c6da7b21d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b7ebcda4aea1211a1a24	2654	5e3639a7c71b3a997aa8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c02c7bd85213c85974d3	2654	ad93cb9499239edb202a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
12d379b20e7b44e84835	2654	53fd09a399c9afcd820e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3a77e2ce6be843d02ab1	2654	1c95850b2ece593eeb02	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ca0df3662ce1e5e53995	2654	0da1a8e4012fbe1989cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7037fed0c59e4c2fc550	2654	83db990ac81dfd8b35c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c957b18e257e1e4d4205	2654	74b1f4902d76a7737fad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6a0d9000a5eed06e2bda	2654	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2fac54b45dcd9764cd8d	2654	6128f2c7422b743e9a52	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
26cc84fe9ba6560faa4c	2654	26f71e77ebb5d7b35f70	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6556edd8953eac56761b	2654	3810d251d2a035b1de85	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
90b294baf0cf8d055475	2654	8899967fd022bd010f96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d6796b335ec054c57c2a	2654	f11f70d30bce575096ec	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2f29ac63597db74c75ad	30	8b094bf3bca9e805dbab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8dedd99a81a811f4554c	30	922a219336536e10a18d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
410732bf1974835a7f41	30	7a6c12bdaf5c6da7b21d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
971b1ad15f0166fd2ff7	30	5e3639a7c71b3a997aa8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5bf96cbe610bbfafa5b3	30	ad93cb9499239edb202a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
82c616f84dfd481054a5	30	53fd09a399c9afcd820e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
552d132ded54d494089b	30	62a6aa3edf4fe2dee2d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c01e536a3e38bc276f92	30	0da1a8e4012fbe1989cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
21b2893b56cbaeae07ad	30	83db990ac81dfd8b35c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
916b6046eb294ca0f874	30	74b1f4902d76a7737fad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
052cefdb9b99f3fa12b8	30	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5282982264c275d14f13	30	6128f2c7422b743e9a52	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aad5da1295b03352a31f	30	26f71e77ebb5d7b35f70	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
10d8344c414e6371235e	30	3810d251d2a035b1de85	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
48feafe4269ffa4663c9	30	8899967fd022bd010f96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3ef82f7b814387311b5e	30	f11f70d30bce575096ec	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5fb95910e720dacdab3b	32	8b094bf3bca9e805dbab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
965b0da4c22b05521f23	32	922a219336536e10a18d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
99b69cc58452389777ed	32	7a6c12bdaf5c6da7b21d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6848d8a70a5b54b2af9e	32	5e3639a7c71b3a997aa8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
24671387799d253848d4	32	ad93cb9499239edb202a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6d6c1f9b811ec6c52199	32	53fd09a399c9afcd820e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
be169ad1244b2cdf0741	32	1c95850b2ece593eeb02	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
90e207cc946821354a1c	32	0da1a8e4012fbe1989cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4210052d0b5551fcfe56	32	83db990ac81dfd8b35c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
530046b2e948496d7ae5	32	74b1f4902d76a7737fad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3a12fc48f103c6e30434	32	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6d7e555a09e297414069	32	6128f2c7422b743e9a52	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6f7bd5f86a0914186994	32	26f71e77ebb5d7b35f70	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e01723180d055c0ce58d	32	3810d251d2a035b1de85	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6cfc75b38fcf9a4e4b41	32	8899967fd022bd010f96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
855a2bf9dde197423536	32	f11f70d30bce575096ec	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
05d16c1d50c039888f07	934	c38a7595c8fe93bb3a4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7ce79d13a3892d83e9ae	934	922a219336536e10a18d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b02795ed0bdf34b36647	934	7a6c12bdaf5c6da7b21d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9b7008fcca7fb6ff21ea	934	5e3639a7c71b3a997aa8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c35f397bc93210674b53	934	ad93cb9499239edb202a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
95a3bcbde8aaee392726	934	53fd09a399c9afcd820e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
126fad5ab3f7e4336a12	934	1c95850b2ece593eeb02	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1df7c1db6175457cfeb6	934	0da1a8e4012fbe1989cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ec1e181d513a7bf2106e	934	83db990ac81dfd8b35c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5bef268d9a85e08ef8f9	934	74b1f4902d76a7737fad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1628592d6f3f5655f0d4	934	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1de6aa96e7611e91cb12	934	6128f2c7422b743e9a52	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8606d4b7eb01ea0f88ff	934	26f71e77ebb5d7b35f70	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d4739eb8c7f53e83882b	934	3810d251d2a035b1de85	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bd97c606f1de1ab8efb7	934	8899967fd022bd010f96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d1893eabec4cb1c7344d	934	f11f70d30bce575096ec	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
53b017c50cc417acb23f	37	c38a7595c8fe93bb3a4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
62d2a6d4170d708b96e9	37	922a219336536e10a18d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
96c42376512eb176f7fa	37	7a6c12bdaf5c6da7b21d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
028031b72e029cd26dd3	37	5e3639a7c71b3a997aa8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7c135bd70f6824e2f3b3	37	ad93cb9499239edb202a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
517ef0e2d2d9f61a34e6	37	53fd09a399c9afcd820e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7720fea363b368830038	37	62a6aa3edf4fe2dee2d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d8434d80535de0ec1866	37	0da1a8e4012fbe1989cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0616280f41ef307475bf	37	83db990ac81dfd8b35c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
da7fe74a2a5e0b4b7e26	37	74b1f4902d76a7737fad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2d2302c12857ea5ee732	37	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5484cf53113aa76fcec9	37	6128f2c7422b743e9a52	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eff3ecad0fdce0707dff	37	26f71e77ebb5d7b35f70	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
78402054dd7df722a451	37	3810d251d2a035b1de85	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2f34bfeda32c9d21ee05	37	8899967fd022bd010f96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
127d79f023703a10f7fd	37	f11f70d30bce575096ec	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
277550caab49c1277a16	1702	c38a7595c8fe93bb3a4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
566dd2b9b37445d69de2	1702	922a219336536e10a18d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3cc1f4ad1c0b10e3bad6	1702	7a6c12bdaf5c6da7b21d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
609d2fef8cfd8aeaf9e3	1702	5e3639a7c71b3a997aa8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a333a18fe6ec3b91d389	1702	ad93cb9499239edb202a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3b712e0153417c8f783a	1702	53fd09a399c9afcd820e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0d40b2ae8f1104e4c4e6	1702	62a6aa3edf4fe2dee2d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3257c02306ef68ce39a3	1702	0da1a8e4012fbe1989cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e253507aa6336e9618c1	1702	83db990ac81dfd8b35c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f6608a8ba574a37aabe4	1702	74b1f4902d76a7737fad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2d85a3242d36f59f6aaf	1702	70d175493126d8dceaf2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eb7663652d95abe472fd	1702	6128f2c7422b743e9a52	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
65e8dc70adc472ec7b01	1702	26f71e77ebb5d7b35f70	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f956e02deca8d8ce0f73	1702	3810d251d2a035b1de85	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
def07881a08599185ac7	1702	8899967fd022bd010f96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1fe89035d8de54088eb7	1702	f11f70d30bce575096ec	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aabf953c1d2b36405916	2141	8b094bf3bca9e805dbab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
71ec695520d4e6299cf4	2141	922a219336536e10a18d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1ffcbeec83432b1d3a7d	2141	7a6c12bdaf5c6da7b21d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8a816f142df41d30f66e	2141	5e3639a7c71b3a997aa8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
65c0947ee77301beaeb2	2141	ad93cb9499239edb202a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1628cc096862074ece81	2141	53fd09a399c9afcd820e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
92b07139c3e3f1650d17	2141	62a6aa3edf4fe2dee2d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
389323521807d2446a5f	2141	0da1a8e4012fbe1989cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
67bc40d4a5ab0e61397f	2141	83db990ac81dfd8b35c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8c521b02d1d4ddd7c85c	2141	74b1f4902d76a7737fad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
66858e9500eae3636935	2141	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
91eb3ec735cd1222dfb6	2141	6128f2c7422b743e9a52	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3d6bcc6dc2de112ef503	2141	26f71e77ebb5d7b35f70	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8172e6301c5dffddd6c5	2141	3810d251d2a035b1de85	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5620f282ef89cb0aaac9	2141	8899967fd022bd010f96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dbc83a7deb6be8741bd2	2141	f11f70d30bce575096ec	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
777d0426a4d63648a5f9	1902	8b094bf3bca9e805dbab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
81754f09c3710c028b95	1902	922a219336536e10a18d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1edcc036a5b89aef2e6c	1902	7a6c12bdaf5c6da7b21d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3b6277f32285a855c5ec	1902	5e3639a7c71b3a997aa8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
909f3226753b7c84b58e	1902	ad93cb9499239edb202a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
59fb0bd231a4911ddc12	1902	53fd09a399c9afcd820e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
069a87e1ae3492bac06b	1902	62a6aa3edf4fe2dee2d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9513fdd95fc2521b2b9c	1902	0da1a8e4012fbe1989cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
597ef297ef0d5a777706	1902	83db990ac81dfd8b35c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d97974894cd004441c93	1902	74b1f4902d76a7737fad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a77854b5c6e4ae0cc7e8	1902	dcfa4362bb9ba014e5b6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
48fb833fa5760bf32508	1902	6128f2c7422b743e9a52	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
67656e1ded9d53919826	1902	26f71e77ebb5d7b35f70	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7e5be35a142b82a99675	1902	3810d251d2a035b1de85	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0d422219b33546b29960	1902	8899967fd022bd010f96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
123f984bda342479dcfb	1902	f11f70d30bce575096ec	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
42d7443ab9de9f1c5e3c	1703	c38a7595c8fe93bb3a4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5ea72e05a9369077b88f	1703	922a219336536e10a18d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2bcceb1c0e340dc7cd5f	1703	7a6c12bdaf5c6da7b21d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
91d5100786ae25e90e33	1703	5e3639a7c71b3a997aa8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cced7b52b77e7ac1becd	1703	ad93cb9499239edb202a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c29760f04caa93302cb5	1703	53fd09a399c9afcd820e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
34d90d5ca472723b5354	1703	6e3f730a3b6e866694e3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ad05903e3a63a1cabe59	1703	0da1a8e4012fbe1989cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eb06b7b57fcdfed81aa0	1703	83db990ac81dfd8b35c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f9a4eed95b2bdfc9cf5a	1703	74b1f4902d76a7737fad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8a63e3cf3c0a2a0346ab	1703	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cef04d0b9ddc6a8f02b7	1703	6128f2c7422b743e9a52	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0e3b4d7a588e77fecbbc	1703	26f71e77ebb5d7b35f70	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f2c9291237c19a71d130	1703	3810d251d2a035b1de85	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a04153f158e08fe48d38	1703	8899967fd022bd010f96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
09d8d7b62c7ea20a3e5e	1703	f11f70d30bce575096ec	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a8213a81c7c0d1e8bb53	2006	8b094bf3bca9e805dbab	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2b04cab5267a779632ad	2006	922a219336536e10a18d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e2d16ec2706ac99f9371	2006	7a6c12bdaf5c6da7b21d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
03d1547fb30ca70649d9	2006	5e3639a7c71b3a997aa8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d0a766669b01c006237e	2006	ad93cb9499239edb202a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2832aeec34c5bf96a556	2006	53fd09a399c9afcd820e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
67a89efe3b93ebb02fda	2006	62a6aa3edf4fe2dee2d0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4ff5b4aa7eda9ae5452b	2006	0da1a8e4012fbe1989cc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5e697c2654c390bfd69b	2006	83db990ac81dfd8b35c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1037d100f2fe80f05ade	2006	74b1f4902d76a7737fad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d7ecf5dbf2625adba5c3	2006	9b82d77926477514bec2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ced111b08bcef0c4d41a	2006	6128f2c7422b743e9a52	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
777bf9b06028c4db09a7	2006	26f71e77ebb5d7b35f70	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a612d4469ea39e301892	2006	3810d251d2a035b1de85	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c8d493c60d58f5b4046c	2006	8899967fd022bd010f96	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6e8fce1b586f92e4a5c4	2006	f11f70d30bce575096ec	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ad9c173dbb973abb6f24	2572	457e892a683f70a78c8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6f311adc543dd3fdf886	2572	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aeef5d87e686ea0f92cf	2572	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
70d7d9ff283e0428cd3c	2572	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
03b9a1af5d03b97e5521	2572	f6eeb477d5f1c37726c6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
787904a58a19d5ec9d09	2572	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b9473812227d69382fc9	2572	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
56794283f0c69161e307	2572	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cb2f5ea0205ac4369dac	2572	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bed1dc112e5b9a02c445	2572	faf6df81a0493135c916	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6b576b4ff29d5aab0092	2572	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
33e91f305f9516671e43	2572	b8232c40c37937eb01d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fa2119e6211c7bb6598e	2572	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e1be3543f60d65eb3050	2572	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d433cc81ab5083f5f5e5	2572	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ae07afb43b60121d50f1	2406	457e892a683f70a78c8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3554dd25affff884ba78	2406	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d0d1309b557e91521710	2406	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
30db928f7f18835b0fcf	2406	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6b117ff2ceb637cf09aa	2406	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
97dff74899f6d6e6f12f	2406	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
221aabd0a5ad60ce67af	2406	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4727c75a4b5009381ff0	2406	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8b60b73d612e95134d72	2406	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
22283ae7b692b9edc28e	2406	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1affc9d931a34da4dbb1	2406	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a2f52144ec6f6b336ac3	2406	b8232c40c37937eb01d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1dc21b8ee540548cbe7c	2406	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e2e253c9cbcfa99e4941	2406	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4d987ad03211528f1690	2406	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
26c59c5f15c51f597d2b	1405	457e892a683f70a78c8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b0a6603daa541da8ddee	1405	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
552ce6e66b653d5be3fc	1405	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7f8557b2d69ea39e20e2	1405	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fd223fab89a54af2802a	1405	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2bdee09de2590e3eb993	1405	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3b09423c7f1807a3b1cb	1405	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0173d982a1220836406f	1405	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1b4e3e47e0cdac876f42	1405	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
92872825287bcb493e2c	1405	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b4dc62f89cae813d7ed5	1405	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b5e0e81a7972b38d4817	1405	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c11dd6734c9e25ca7e4b	1405	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6a6c44b379414c0f5512	1405	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2497d50d47e89638a138	1405	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3a281eebd7d5124da46a	68	457e892a683f70a78c8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7b983cdf0f894ed5cfd7	68	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bb3aef016e5d00839015	68	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c0b1556881174c928a3e	68	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
db14fca2c3b334164955	68	f6eeb477d5f1c37726c6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
33489751835cd0d27b1c	68	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d784e1c39cb14ca1407b	68	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0cf03222070bab6d9b97	68	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d01e7b52c58003c9a769	68	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1602fd7072ad63f30cea	68	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
09bcff11d9f7add1cf52	68	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e4f04d7d9ab6a47b48d4	68	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
838f69ed3c1388405391	68	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a23fcb32e65578e848c1	68	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
84a0ff2c396e15d89927	68	b96a13884d9cb2f18ad3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
134250a5981671241615	1187	9a64dc0d5caa2750bf59	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f31711a72ad047922d13	1187	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e96f0b0e1a34d6f480d1	1187	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
583db075610c16bf167a	1187	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6e1b8bb26cfb1ac6c950	1187	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
07750e4189df9579f7fc	1187	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
73e8478598dfa5f94dee	1187	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6efefcfaa6f7237d735b	1187	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
24ea72e603d75af4c640	1187	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ce320a3a24bd71c07788	1187	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
274d58aaab01ffaf2a0c	1187	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d08c784bef5fb3d5b272	1187	b8232c40c37937eb01d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7cdd5e825beea29d9762	1187	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3423d9f740d1f9dc5539	1187	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
92c74bb00ecb823b2dd4	1187	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
80f8a3e08278ea305960	1943	9a64dc0d5caa2750bf59	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
98292a544f840b2beb04	1943	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f7b4cff59bdb794bf1a6	1943	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
252a28a12fcfebc5f0be	1943	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f89a38a290287af5e46a	1943	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
acd2415e0c5e6da53779	1943	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e5e1bdf334b5350c4edd	1943	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
71317347f2937769c66d	1943	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4182b4ebdc7e78e45ba3	1943	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
558f214c5449fff6a8f9	1943	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
054803914f807f8bcdd9	1943	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2624d575981ae59c02cc	1943	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
989dfb8748a3b6060d60	1943	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c5507a91e6656a577ceb	1943	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ba82bd278b08c0140a09	1943	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c36d5cf9f3dffef21823	81	457e892a683f70a78c8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7437be9ebafab210f70d	81	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
880b36fc0049428b1ef9	81	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
61110be3c459bf9c56f3	81	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a6e0b8af00b40059b473	81	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
25714a9faa8e1e9f0f53	81	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d5cc11344b0653cb68ca	81	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
972cfee85be87e67ba61	81	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b557ad135e687859e066	81	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c1dfbebf8b4aefb1d865	81	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ac904c497246d3e87600	81	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
26ff4eb9b989f4318e99	81	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5bd39bcccdfb77d2602b	81	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0d8e5e4f23d97baf2c83	81	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2851cdcea60192c78c89	81	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eb659ba8f9ca41ab78f1	1540	457e892a683f70a78c8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
68f3e77ee4b8b3a9a883	1540	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
20a79eafcbbfaad33a11	1540	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f7358f4f1ea2b200aa1a	1540	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dfa6ec9fcbd1570ed45b	1540	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e2ae99643080f251349b	1540	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d4854af5decfe26252bc	1540	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d9d8e885658f5a4c960e	1540	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
426f5ee634bec2bb4e26	1540	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d5b35ecbeaba63e6c800	1540	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d3fff5987c035845de3b	1540	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3563eff6fb3d43c7d18e	1540	b8232c40c37937eb01d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0a16d2d20e3adb8ca6be	1540	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2db7c42687faab2ccb66	1540	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c47db6c1baa9357f6e3a	1540	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
75fefd45cba4f110563e	1538	9a64dc0d5caa2750bf59	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9b58ef502fcdfcc8c412	1538	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5ba425a478f3d8e422d2	1538	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
61ad09ffb36c69cde4cb	1538	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7005bf2b02ceb573b7d0	1538	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2149952913a4f208c1db	1538	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3acaab4b36411848a28c	1538	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
87ff357e2652d2e6285f	1538	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
128ddb88e9f26402161a	1538	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
52365d06e5daf3279fbd	1538	faf6df81a0493135c916	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ca1076005ea80b85d05f	1538	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
34ea1492d983fcebb501	1538	b8232c40c37937eb01d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1989ec4464ab73ea96dd	1538	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f13cfbc70904075d376b	1538	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1e9e342a056dca8a81a4	1538	b96a13884d9cb2f18ad3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a37ceb2f6e1cbec094ae	2378	9a64dc0d5caa2750bf59	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
02708bdd3f2d88ecde95	2378	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
663d4eea872ce4d4338c	2378	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
33243bca783d186c9def	2378	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
15dfe6d802026bdc80ad	2378	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7470d7570c9243737799	2378	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f6f4121424644c7c4ab0	2378	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e501a0f2fa191a384ac6	2378	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9b85c286feaed5601c0c	2378	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5f3c5e2b99fa8fab203e	2378	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f27e30181cc72e4a53ef	2378	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6331f70e930c6073802e	2378	b8232c40c37937eb01d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
94787e63044689d1f079	2378	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1841498df76e9ce36058	2378	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b093462b7b4d40790fea	2378	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8a921e35f9b02bce910e	2320	457e892a683f70a78c8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4f4a8a71949ffbddca09	2320	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
51a3a8f4f9b48e5a5c7b	2320	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
826516bc34a2f77b43a6	2320	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
855e2966cd8f8ebbd3cc	2320	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cd5c4f563092fd4ef1f7	2320	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8539005ad8217fd9cf34	2320	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
05ce42ebf5379d728ecf	2320	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8c61df9cd6c36e4518b3	2320	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2d0c04276fd531e40449	2320	faf6df81a0493135c916	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5ee784efffe5105a06d5	2320	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2f8be85942052556619c	2320	b8232c40c37937eb01d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7f106e4d8a0daa32cceb	2320	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e21784b648b25cc4f1d4	2320	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
59d28bcaab2e907eead7	2320	b96a13884d9cb2f18ad3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1738b3d304da889deb7d	2240	457e892a683f70a78c8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ce53f100867053e00e63	2240	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
90b2c91af42d65c8d1ac	2240	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a47f85be499e39cca3f2	2240	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9ea6f8d2d948c37c219d	2240	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9b2170566e8e8b668a62	2240	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f453e3798172d2cf087e	2240	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aecb00543c2f43bf79f1	2240	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
efbb0c7f85a49701fb62	2240	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2b6169fadc1775e2f6b1	2240	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
db9e01f692c01a53ce46	2240	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d594c7c33e577c88748d	2240	b8232c40c37937eb01d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d1e6bb7801234bd01a05	2240	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a30f392a4cbe6de27c95	2240	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d97ac7884c207b19dee4	2240	b96a13884d9cb2f18ad3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e44e9a7dc95a8dbdd543	1764	9a64dc0d5caa2750bf59	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3c0557151dffa4522925	1764	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bb054bada0d2e95761df	1764	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
48472b9f91f8d433c245	1764	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b705f55dbee4edd50447	1764	f6eeb477d5f1c37726c6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
624bdf1a8305e57f1053	1764	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0c2604de4761756802cd	1764	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4ccc3fef358f9bd9e55a	1764	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b4957a97bce8a35156be	1764	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6bc9a81c5609152199c6	1764	faf6df81a0493135c916	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
334094acdd7e68784c40	1764	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6a8603ae6106fb8205e7	1764	b8232c40c37937eb01d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
86c7abcd4d44c3e9e97f	1764	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2ae27c533ba2b4b721ea	1764	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8eddf2500626e02716b5	1764	b96a13884d9cb2f18ad3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
48e454d099c511ab0f20	2243	9a64dc0d5caa2750bf59	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c8d3bd150d05325f1582	2243	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
df76a450119efc411554	2243	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c67de7363ba8c5f60da6	2243	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a4ca375f25be95134e96	2243	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
99ea44ca6e686e1ae0b2	2243	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
36f51c909dfc1cf5a53a	2243	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
33e568fffd1d8de89152	2243	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b6704f478befb2b4fa88	2243	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7d8708d98979d484e8f3	2243	faf6df81a0493135c916	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cc5f0b9cdb1599e0709f	2243	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8518d297e8c01c96ea7f	2243	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
16cc34872960af814ef5	2243	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f22abc4bed3ca94a3321	2243	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b4379fc82897e33d3213	2243	b96a13884d9cb2f18ad3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9b51fd2a28f077016d43	2576	9a64dc0d5caa2750bf59	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d8ca5b9ce99be847900c	2576	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
41cda57ec068c2a8f653	2576	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
731680999ca318e1e493	2576	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9824060cd2e96fe6f1d1	2576	f6eeb477d5f1c37726c6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5a9c65111af251442eda	2576	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b0800d151acea38b179f	2576	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
702f46ce314f88cdb0e9	2576	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3d38c05a0367642a02c6	2576	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f40111616264cff6852c	2576	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2a1df338db56aaeca2e9	2576	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
821bd7c736194d3e70bf	2576	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
624e04ae79d6f45cd641	2576	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e33cc53e8b024fda5ae5	2576	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c35c27787d61d4d0667e	2576	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4ac9824a9dbd3640fb83	100	457e892a683f70a78c8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
71715cdf0534c0cf75c5	100	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e385e3b4ff989ae579e9	100	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6dd13c35e6b519b4658a	100	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
45e6b00fe1cc81cfaa4e	100	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
323717876b72bd594164	100	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
90168df906b010b1153e	100	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1aca7f47b8953c66d844	100	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c2311101769126071de5	100	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4e92053eacaf305d67cc	100	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2d2d1b5f4904177d90b7	100	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
763c6fd7cccd4c58a092	100	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c9073865b7ed32c6ef10	100	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ad74370d2f88d2f9d6de	100	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d21c24dfe87b00a84cfa	100	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1e2d4b9d0087c06e34c8	2577	9a64dc0d5caa2750bf59	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a77745c2fc4013493d8d	2577	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8d4dfbf254e03d8f491a	2577	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f4b21f6e26bd002be343	2577	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0eb1be8a6fc4e804ba08	2577	f6eeb477d5f1c37726c6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d927866be8e29dbf44ee	2577	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
332d21b91fa0088b8799	2577	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cfc8f90b1a92e647f53c	2577	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c1169efffbfed6bef5a2	2577	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
01b5db6366891f3c1783	2577	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
035b6d11e75dd2c05c20	2577	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3f630e6868927f33c5f9	2577	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8c48417f90119c621cb7	2577	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b9e8aa0f928a7150b21c	2577	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
da0f053996768c397d6c	2577	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
04684821083d3768197c	1191	9a64dc0d5caa2750bf59	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b7f403b484baf41a1d08	1191	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
81496f6876c39224194f	1191	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2f8649f0570d986d7690	1191	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
96aed003ef86e3441e6a	1191	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3897c7509fcd945593e0	1191	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
348fcb46fc52627e881f	1191	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
965bdedd536dad81ac49	1191	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
87a27d9416dffd1a7149	1191	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0a1acc2dc7824b1b483f	1191	faf6df81a0493135c916	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fe6ccfd89ced71167661	1191	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
120b6458fbdfef49c265	1191	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f03826e3bbf9edec59ca	1191	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c7f358cc0430135471f9	1191	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
00962202d5fe4687af55	1191	b96a13884d9cb2f18ad3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
36163c874339fe5ef84b	1198	9a64dc0d5caa2750bf59	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f059bed0a906decabd2f	1198	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
acf53f4d26964a1ff523	1198	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b6bb296e595a0cfc5a83	1198	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9ffcf43df356c8615592	1198	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9cc243a38464088d723e	1198	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
326bd26a95cbf2b9bc20	1198	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
01565d7c92267074122a	1198	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
63896e2600ca0584e7eb	1198	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e815ccf83ded0467107b	1198	faf6df81a0493135c916	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c5b8c226a3f7cf3fa0a6	1198	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
76a5e266b83eb19a2b9b	1198	b8232c40c37937eb01d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e9185ec4a60986ef2bb9	1198	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cbfe30f874577991fd78	1198	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
54af2ada2483e7a3326a	1198	b96a13884d9cb2f18ad3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0d34da8e19bc652fb57d	2401	457e892a683f70a78c8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
49dc039e9a3a18a07a8a	2401	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ad2f65023ebb2d6eb61d	2401	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9e26913ca606660fb995	2401	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
11b934c0ddbe0d12047b	2401	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
df61b8e9c63ad6f5eb5a	2401	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2955a4105259cc8d2904	2401	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b455b58d1fb34ccf93b8	2401	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c14e323eebcd63ccefb8	2401	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5624849523ffb248975d	2401	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e4632a5d95bd87658992	2401	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
401f58135d2388b3b2f2	2401	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7dfab23f211d47577bf5	2401	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
92ffe2768e5cd5abea41	2401	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4f25d759593048d8c30e	2401	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c2649693624245c8351b	2147	9a64dc0d5caa2750bf59	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a21426c6cfd2d99c8810	2147	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b1d12d64c83d42f32cdc	2147	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9db86a76bce85e6c8adc	2147	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
43e42fbc7f119cc86ae8	2147	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7118143a6f98621788cf	2147	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bd123e445f9780fa3256	2147	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bd7b1e9484bc0c7fcb4a	2147	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6764e9530fe921e8e2a4	2147	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aa2c82b1d50924ff9249	2147	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d0d019f1522269ac2f42	2147	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d0e50594ebde27a4f2df	2147	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
84bde41269c512b02785	2147	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d3f34b1f4771cc458f7e	2147	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0e4db24984040fce9862	2147	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
526db5ad04cdb70a1ad8	1194	457e892a683f70a78c8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f15ba42cdee6a31ec1a7	1194	471952423957373168f7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
985dda5dd1fe3c4107f2	1194	0371cd4bdfd4b24f8cbe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
65541bca706ca5d9ee6d	1194	0b701012c2790c7c4735	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2e4657748e37f06ce795	1194	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
83bd1779ff1d05e52dd2	1194	1f2141b07765892cb67a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
512e5834a89a8cff1d15	1194	7e5ffa821f50c0bc8f8a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fcd75b62bb2546d29287	1194	862d7a9266b086d90ae6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c2a66cde61f7b768dbf7	1194	117e47621ce29f1a2bc4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4434119444b81e59def1	1194	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6738d6584354066cec30	1194	877d41f9b7669c490ac5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d14725f62813d4aad022	1194	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
54349b0c6d097a9773ac	1194	f37759490150fb4b8c33	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1a89107c7e9f792b09da	1194	97d9dad9b451c17681bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
45337e84e760ff29a701	1194	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d6c2a6daaeb007311ba3	1129	99e6504888bd1d9f9986	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7a8c731b91755a539ed6	1129	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
63da240ef2502a02ed99	1129	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9a69ffe3ae931e8065bc	1129	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0807bc0c744576631661	1129	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a3442129744eb8cbafe9	1129	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8b4c812286d0528a97d7	1129	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b74887df84321a5db9b2	1129	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
932840ca9dc90c8085a5	1129	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ad497b42561af42f24b3	1129	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
52f7c6de31006b7bd38c	1129	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
52788211f05cfd6e9bfe	1129	b8232c40c37937eb01d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
431addcb46dc0b69efa6	1129	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eb2567d5100a55ee0df8	1129	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8559ea033b56b38cbb42	1129	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
07a251d9b1b4a6649f73	2573	c90d77745ab6d0e48537	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bfa72a726f76063cfda0	2573	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e127713beab0c6a6adc3	2573	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2653b95e2eda7730e1f0	2573	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b028eabe519272f251b0	2573	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9485c8546677949be908	2573	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0c1d4db1bf709a30c028	2573	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
92187e1ada34197c3724	2573	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0684ef4ca56a4c0b2c38	2573	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3e81a648af40d3f5d136	2573	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
003a1e7a35118ffb4227	2573	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c6f0a9b6711a1a43e8ab	2573	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b067a282ac65f047c9a4	2573	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b45b742698d90c20f8a1	2573	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
152a1d768912310244a9	2573	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a26a214b507b86d6e09e	1088	99e6504888bd1d9f9986	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee734b10b4ce41f61920	1088	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fd98afa351944bebd02d	1088	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6cf44055c40a2dcb54d0	1088	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e634942ba66d18472efb	1088	f6eeb477d5f1c37726c6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dd50286178e6912db9e0	1088	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
04afddc53493bf30ff61	1088	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1d6a69fe03273e6706d8	1088	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ba506b277d37407082a0	1088	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b73d8b0bcf7c036f40bd	1088	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e510080680502ca875b0	1088	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
00fbf69f44acb72ec070	1088	b8232c40c37937eb01d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
65bbae24b3316ea37041	1088	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e9009bec46e2a763bd6d	1088	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
68f87248b18e0ddc0f40	1088	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a40a1c4c0e8e9fee90dc	113	c90d77745ab6d0e48537	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c51b8b52b3e2b91982be	113	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bc75acd684a04faac206	113	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cd8f7efd8870194e1970	113	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4798dbe133df3db0a9ab	113	f6eeb477d5f1c37726c6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b68635e195bceb3ec864	113	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2c51f817fe024d2da915	113	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
326b426fef387ccd5d48	113	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a88c22c3cfb476950e42	113	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e08b39723d3294221260	113	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2c8374aca289d02ce5b1	113	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
70b02c479180655a0d84	113	b8232c40c37937eb01d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7f354f9ab48746f284ea	113	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
05a02883c1af2e4eccb0	113	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
62f6c6c9bc9e18b4ceb8	113	b96a13884d9cb2f18ad3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
08ab1c9f338d8ba846ce	2410	99e6504888bd1d9f9986	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
449829a47d33dca609c5	2410	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d073442a4deb324f6879	2410	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
97295e6da2821ee71b7b	2410	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dd6a3c6de916f9df1ad5	2410	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c9b2ff05afbd2dca536d	2410	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
afed4a93331603a30a40	2410	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
69916c098263804cdcaf	2410	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f896c3193599663437ad	2410	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f9e9f045aec5e49f798a	2410	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0d5b0345dff051f709af	2410	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
47465064994401a3c329	2410	b8232c40c37937eb01d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
946ba15505895f01afff	2410	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c1bb7ec16d499b2b7378	2410	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a76db8d04c0d7f5850b4	2410	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
488659ad314432482a7f	1110	99e6504888bd1d9f9986	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7fc611f2fb76cd7e6fda	1110	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9d9fcbb0ef736aa7a490	1110	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3c3fe608d4f6005253fe	1110	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dcb4a1fa433c7ecfd5f6	1110	f6eeb477d5f1c37726c6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1b8f080817e7acda9c43	1110	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f8106c8e048433fb3c64	1110	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
124846520297d700a1ca	1110	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
beba0a9410e84bf17c37	1110	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
948b6af28dd45d257210	1110	faf6df81a0493135c916	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
36acf673be8c5c1e777f	1110	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5dab3ff64fdd7c3e59be	1110	b8232c40c37937eb01d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4e9b12db8ac775c1afba	1110	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c4d880e82cb38e2f07d7	1110	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d13f7214678fdf2ba5cc	1110	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
650dbdf7b3ee60cd4eda	1944	99e6504888bd1d9f9986	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4fd89e987f333ef0411a	1944	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b1dfd70ed08c46258a24	1944	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
89c725de4079cc83e163	1944	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e0dffcd0ee588cd734c6	1944	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9ea84048d03422876255	1944	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2a0ced2842ff7afe81a4	1944	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a2029d25dd5f0959f2b7	1944	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cef393859c2c5c6750c2	1944	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ea19ff576398c8fb997d	1944	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9955eb696f49757c06bf	1944	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a00fd2dae6d751376b09	1944	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3909646c64f61bf17c9b	1944	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0bf4aaca252cb9b827e3	1944	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a562e5cf6c8e43a7dbc4	1944	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
43fa774138e7d25e581b	77	99e6504888bd1d9f9986	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
277c66300953c14af8e5	77	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0380f0a606772901ecf0	77	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c0a3148b5676b4963f76	77	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a2b2287381bcead332bb	77	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bd55dd2bb2946adc102c	77	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
40eb3cf867178ad2401d	77	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
41760731ea972e8567da	77	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e8c58a83e85e36643911	77	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d73a1aa6092ffd811dee	77	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f643e99d1a1f927fe7a8	77	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cb56379f0affcb0c1239	77	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cc57249d19c97d7dd626	77	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
69a206099119cc4d9554	77	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9d7ce37dc79fc96e61d7	77	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9aef1c54e3730c08c877	1560	c90d77745ab6d0e48537	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5d4abca9e3bfeccb0f09	1560	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b3b006e6b480e9486fc3	1560	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bcf67e46014f01d7ab64	1560	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
830e0693f4e91bb012c0	1560	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0f6bd21e89e58c106264	1560	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
51d4040af79ee256be87	1560	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b99b1f23a2cb7d340dd0	1560	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d4db024a03b7657793ca	1560	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
58c0a19360015366ac88	1560	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c7d8aed7796c3b840fcd	1560	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
94986239d12ae8c8af7c	1560	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f1cb35c8fef2e9a431f0	1560	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f9a14cdbd71458068a15	1560	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ade20e04e441e5e584d2	1560	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c1d43ac8b6edec925dc0	82	c90d77745ab6d0e48537	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
198fcac3cec61949fb43	82	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0eb368596d0eba002cd4	82	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6e57fd492dd93607d4ff	82	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d305b70fa1f1ac0e7b2a	82	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dab1b6583f3d4e986f04	82	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
83b6394c439fc86d0878	82	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
11bbd03de605bf3a1803	82	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7134ebd2dec486c9ef2e	82	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
59edefd6ffe5af88f4a5	82	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3ac8e88a397e47915d7a	82	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b92f64faa5dec826f130	82	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
805adf76356b5e85c990	82	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8dac633047290eaed561	82	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
826d75e2c84cff8c1ca6	82	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1c0d71a005d60beb8db0	2604	99e6504888bd1d9f9986	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e240bf8d77732a0a7002	2604	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
48bf67ed46622da2ea5c	2604	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
76767aa9fa2c28fe7c81	2604	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
40589cd1ba871fbb294b	2604	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fbaa8dd505aea95680d1	2604	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e4d027e2ae21954426bf	2604	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b2d923ce9c31bfb52878	2604	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8ae2941120b4c1f19737	2604	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
559d4d1d4ca6e60f4d30	2604	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
76e953dd81d710100612	2604	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f1d44e50f3c5a7fe3a5c	2604	b8232c40c37937eb01d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6c2ebc8275875ae1b036	2604	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
936bbd01f23812a80544	2604	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
010db4dadb34f04d10a4	2604	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0495d7daf63214dfa687	1197	99e6504888bd1d9f9986	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7ecbd7746bdf2307b4da	1197	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b191fd63711478b6ee86	1197	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e2a682749593caa17dd4	1197	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7210510130e03b7fd824	1197	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ae0ccd8de2f46c550aea	1197	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3b1f1b9ff15cd1718a2b	1197	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c20d6b5ac7108a6f299f	1197	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cfa3ce3d663d0c96c292	1197	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9fa06dab50c767823c2c	1197	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
25fff6e53c15b88094c9	1197	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3609b9a65f630ac33219	1197	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7e9fcd8a01bea5e903a3	1197	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0b8eedb1a61230a5d15d	1197	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8ce650b56cf799b47e37	1197	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a82a3de8deab0baa535e	2211	c90d77745ab6d0e48537	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1c04711f8521c506577e	2211	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6bef7daa209a5e60863b	2211	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
524180e50d2e9914acee	2211	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1b0bd3efa1db4469a695	2211	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0f6bd445c184712d3242	2211	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4cce03e2eccfbccb3b2b	2211	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cf37a52f949819194236	2211	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cafc6b58afd34c4771ce	2211	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a8f2b74b05eb34f5b878	2211	faf6df81a0493135c916	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c31bbda1649a24c2909e	2211	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
70ac640fa6026cb20e03	2211	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cc6a7b74907b0b40c326	2211	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e3fc148190b11a65e189	2211	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
901cd0c25c0738f069f2	2211	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e9f07aef0e47c843199a	2429	c90d77745ab6d0e48537	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3262e1888fd02ffef3a6	2429	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
85baa96f9c58389d9972	2429	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
da3239070c8a88fd2f66	2429	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9100cff0d0950709d12c	2429	f6eeb477d5f1c37726c6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2d877820d9f02447f75b	2429	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9788bdcd77fa174d7806	2429	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eba7ebb29e046fe7e869	2429	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5e229fe60dc804c5e806	2429	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
224818aa51a56d2086d0	2429	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ae0d482e002555ebe54e	2429	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
07333c097b92b813d760	2429	b8232c40c37937eb01d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8f57dd7f7bc38640da39	2429	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
90fde14371b1b5597d5c	2429	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3a790bd3ad1852321dcb	2429	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5c6b75bbcb2d1fc49cc9	1460	c90d77745ab6d0e48537	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1b91aefdb61d8fceb486	1460	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1da06141dd7d42fb0030	1460	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
530e7a2497e86069c635	1460	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e64a6a3fdfd3449bb572	1460	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e64b71df590f8abfec93	1460	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2b87f1d0699d44f8ec34	1460	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee925cce79f607e775c4	1460	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2e03ffd609a75ca1e331	1460	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
767d55a0f32323dd1e2c	1460	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7167da746f9552bfc80d	1460	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ef7bec7958302dbb174f	1460	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1886fec1214d9cce986f	1460	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
783fe5fe29beee7b2bf5	1460	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fdedc1a9896ee6e34b56	1460	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4ebd5a143b0c5ee28265	1717	c90d77745ab6d0e48537	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cb65ab427f976b02c509	1717	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b989b58929f084b2c58b	1717	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b12a052ce408520b8167	1717	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0134d71ef5f693f5d8c0	1717	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
37fa691cc2ee4a007538	1717	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
51cc65e13b24ba40d2b9	1717	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
172ad8fe3a0257a0e654	1717	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4a5ecd5cc6fdcc51422d	1717	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
025eb6e0f87dfce3328f	1717	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4fc1e1342623fa46e191	1717	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b9ffb744662871e5248a	1717	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7bfba3afb87139915506	1717	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9b3c840e8e8845528860	1717	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ff6560236431717ebe94	1717	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cb6e40ef6b8a090553f5	2578	c90d77745ab6d0e48537	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ace3d57f99712acdfd74	2578	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e3eb58400f937b0101c3	2578	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bbdc78326c5c2d7240ac	2578	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7a117393dba9aa168d6b	2578	f6eeb477d5f1c37726c6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4b8fd53a9f7e7d13dc2b	2578	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
113fa21841f228ca1b59	2578	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
310532e87d599d768767	2578	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0b96b54b131b9c5dcd0d	2578	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a5f3979fad74219f5848	2578	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f695b258a267385626d2	2578	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7be13d731cc95c729c61	2578	b8232c40c37937eb01d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3ca057b1819f2068c8d8	2578	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
11a4ed843300cbb1e8f8	2578	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
29d5bc27cfe8b54dda1a	2578	b96a13884d9cb2f18ad3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a83cb0c7702e0ca71332	1572	c90d77745ab6d0e48537	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
34f24d8ca724cc4d0f12	1572	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
802b6a41ada9be23eece	1572	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fa034d7f9933da92adb6	1572	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8e706cd7b66b597fa588	1572	f6eeb477d5f1c37726c6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b2b9878822b26489c105	1572	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0c7804dd186d4c4992cc	1572	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8c0dbad46926d4a746f9	1572	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
90d7fb52274adc4a0960	1572	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
321b8157841690e191a8	1572	faf6df81a0493135c916	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c0c7fdf1ec4646a34f29	1572	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
42e5cf70cd57c9db3fa0	1572	b8232c40c37937eb01d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a13f7e4c7e1a8cf0101c	1572	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d507b7002bb0ec324001	1572	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
adb1aab015bcdcf4400d	1572	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f087712460336f23fe80	1806	99e6504888bd1d9f9986	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
662ae361bda17a3ad95c	1806	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a99b5b9a90980f8eb8cf	1806	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d279f5068ed8a3de6117	1806	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a1ab52c1337a1108e1d1	1806	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
df999bc2c759628f5a3d	1806	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
01f7af1217145b6b627a	1806	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0bcae38d9b724f7737b1	1806	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dbf32b50ae4c281130f1	1806	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8af8451430eaead50d70	1806	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dec056c5396b5efcd188	1806	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3c113a0a6984bc88964a	1806	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
424e321bf2add3c166ab	1806	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8a97f1b3abc915adf75e	1806	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
164ac781ff91bd4cfcac	1806	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aad76350dfaca5a7ccd5	115	c90d77745ab6d0e48537	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7225f5015779da4fe13f	115	bd6be185e5e6e11ada1b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
98ebd7846405a9a9e997	115	76d5185f429653ed2e87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
79a76ade85d07dfd10ec	115	2f7f607a87c99c6722e9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
93967c0c895fa7b85189	115	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7879256e0f6ba4e00b69	115	dd1d596229411ff299e5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bcad546ab7c5d3704978	115	04dd4b67bcf35901231b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
53bd2f28a4308216482a	115	74f2bc4a5877a6277d7c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aa98266e0c12e3ee661c	115	649153b9a1951bb4bb9a	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b21777f67d6201468629	115	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f7e95e034419edae3227	115	bb3d38b9bf77a8e95b79	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c9c3ee7b00125593c23f	115	ed07ec8fa564381ea093	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
058bc047abac9127004c	115	dd3c19d41c8dd7d27682	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
45152a817238dd72c520	115	431824f1591821900428	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
62ef50a30555c6be17ac	115	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
955e55f0e6fe5a523163	2223	8427cb0e609c6395011d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
55fca7a9f8ba06a2ea52	2223	cf5177242c1034802d8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
88df7a2154e3c76896d2	2223	ec964e12ec298f154906	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5aa638f0484882d92f5c	2223	7f18686d985523c66aa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9263a95b2230c19bba48	2223	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
52a1efe69e6f315f8c29	2223	b232a97de45a99e40e65	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c54c1164dea12c45085a	2223	50539e1ab2fe53f5df15	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f7ab8033e62dff269bf2	2223	86c46654bedeb5a02591	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7bfc7455ab89ecbcf1ac	2223	0f692345829560ccc163	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1df12117101a1d273f59	2223	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
657c2800d2a5786e3553	2223	e48a2537426fd384e04b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6efd2048d7531aede1d8	2223	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3369c674da40451fc9e3	2223	0a30a4d85b028245ef62	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
20c1d9f0e6d664f959ce	2223	028f24b1257fd6f8b5af	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a5e6c86d23a4645e81c0	2223	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
496ded46aec1ba30f870	1113	8427cb0e609c6395011d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
741eb4048d160d046a03	1113	cf5177242c1034802d8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e4835cba2da6408a5f90	1113	ec964e12ec298f154906	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e86020384f3d509e2032	1113	7f18686d985523c66aa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
40af21d5272dc8fb268a	1113	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f58e7e9693de31832e9d	1113	b232a97de45a99e40e65	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e3ba8ed1300caf95912d	1113	50539e1ab2fe53f5df15	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4acf34ecae340fa84f81	1113	86c46654bedeb5a02591	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
38cd7b977c474848f755	1113	0f692345829560ccc163	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4b70481ffc7daa75daf2	1113	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3fcadf4d097f206e9773	1113	e48a2537426fd384e04b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8d5452e1ba6aa3ce5d1b	1113	3ce1e9bfc95928b110dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6392071957b7826b0b48	1113	0a30a4d85b028245ef62	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
52a4552b95097dd0a5a0	1113	028f24b1257fd6f8b5af	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3baaf07adb8a5b2d2285	1113	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ff72d6f070a6caa39e09	2574	8427cb0e609c6395011d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
32093b6e8aa7f7838325	2574	cf5177242c1034802d8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
25074221d2c89dc24fa2	2574	ec964e12ec298f154906	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
52c7fb0ad3b76975f8cc	2574	7f18686d985523c66aa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
45e598baa9909fdd0d13	2574	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
768e1d1e7c9923716a90	2574	b232a97de45a99e40e65	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bb013f25ac6650e0963b	2574	50539e1ab2fe53f5df15	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a3f71db92103a3cfb824	2574	86c46654bedeb5a02591	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9d4986c36ebd5f8c2805	2574	0f692345829560ccc163	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d5fcf2e3434a1b2518be	2574	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
709ac2bb8a2b13f23f16	2574	e48a2537426fd384e04b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6dead939f35e732e92a5	2574	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
49b805194abc76277ba7	2574	0a30a4d85b028245ef62	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
56af15291ad8d80e78ff	2574	028f24b1257fd6f8b5af	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7ab081680271a888d529	2574	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f0d3b33377c839e0f10b	1709	8427cb0e609c6395011d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c96cbc7e14c3ec20ff32	1709	cf5177242c1034802d8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
23e0d2a8de4a42fd7299	1709	ec964e12ec298f154906	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9597fc69ca138917e50c	1709	7f18686d985523c66aa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
05685c87eec7dd095f63	1709	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
32ee8b834402d2d9d420	1709	b232a97de45a99e40e65	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4d77b04415f28e14f17c	1709	50539e1ab2fe53f5df15	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a4c173223fdcf788915f	1709	86c46654bedeb5a02591	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3a5625f8e567a008d857	1709	0f692345829560ccc163	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c37e21912cd7b187b381	1709	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5b871a97da5d9857302a	1709	e48a2537426fd384e04b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b0585b4cb3411453bd11	1709	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3b4940b047c3e405c507	1709	0a30a4d85b028245ef62	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
076898e0130b0f89267c	1709	028f24b1257fd6f8b5af	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1462acbe18f6b8c11877	1709	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7a0cd60f797bd41d2dcf	1710	8427cb0e609c6395011d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ab586a6991bbae29f48f	1710	cf5177242c1034802d8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
189f074b1629949c8462	1710	ec964e12ec298f154906	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
82e784b30d1a140598fd	1710	7f18686d985523c66aa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a3582cc03b094c96865f	1710	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
53121a0536886a83050c	1710	b232a97de45a99e40e65	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
866678b345b8b54315a6	1710	50539e1ab2fe53f5df15	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f70e0db0c50b504e12cc	1710	86c46654bedeb5a02591	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2a16e63061bfb4cf280b	1710	0f692345829560ccc163	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6f3bdb25795f9ee64725	1710	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
044c68e48bc9ed845467	1710	e48a2537426fd384e04b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3979185050ba5892162a	1710	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6458de402090d70311bb	1710	0a30a4d85b028245ef62	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
97c787e9a4b94d3323d2	1710	028f24b1257fd6f8b5af	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
923a25660090f1112e82	1710	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
defa4a297d35cc9213bb	1908	13cf5f07e77de99dcc30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e37b762123a0db2c5289	1908	cf5177242c1034802d8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3394178af5fe99c3af9f	1908	ec964e12ec298f154906	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
868263e5c36e3b37b6b5	1908	7f18686d985523c66aa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f19d927bf100d27fe6ba	1908	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7be6fedeeee9c0675676	1908	b232a97de45a99e40e65	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
138d0e8d1e11db546f87	1908	50539e1ab2fe53f5df15	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4992119ba4e8f326d386	1908	86c46654bedeb5a02591	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8cc61d2c586b0599fedf	1908	0f692345829560ccc163	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e8bdbcbb11d94e855f72	1908	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d4454662530425dba580	1908	e48a2537426fd384e04b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
08f84cfcf923cc7b82ec	1908	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1ad45493aef8c0cd30d5	1908	0a30a4d85b028245ef62	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2b8cdf16f395ff4af6e0	1908	028f24b1257fd6f8b5af	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4c6c9d6be33921989bfd	1908	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e918d65de655f90a61a7	2293	13cf5f07e77de99dcc30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
91c3b291bab69d4bc14d	2293	cf5177242c1034802d8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
467fd285d88e3967f660	2293	ec964e12ec298f154906	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8f6b93af2065b28d86ef	2293	7f18686d985523c66aa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fa9fe2d3393acf14b834	2293	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
71204c78d69e1aaead80	2293	b232a97de45a99e40e65	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a360fcc3ae82718f702e	2293	50539e1ab2fe53f5df15	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
093b044076f3d47d259d	2293	86c46654bedeb5a02591	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fc4a38ef3e49e1f5a8a9	2293	0f692345829560ccc163	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a1d260389b7431405c23	2293	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eb59d2b68303c10df5da	2293	e48a2537426fd384e04b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4120354fbc3b08c5c8ce	2293	3ce1e9bfc95928b110dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2ea0b9c7c39f91f5c4b2	2293	0a30a4d85b028245ef62	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
34e8804c72e4797f7afd	2293	028f24b1257fd6f8b5af	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f90a015dfff9b24934a4	2293	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
289a3966c82af7a2e3a7	1905	8427cb0e609c6395011d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
43a1927ad353377a4b25	1905	cf5177242c1034802d8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fc0a3021581e566bec67	1905	ec964e12ec298f154906	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ec7596ab3b6dbde7d267	1905	7f18686d985523c66aa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
370b42f4861bc449f365	1905	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2e83e55821b3f7e854c4	1905	b232a97de45a99e40e65	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eff131a9a5b0a1d78545	1905	50539e1ab2fe53f5df15	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8b1e1594e5d56cc761bc	1905	86c46654bedeb5a02591	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
36f4b4dcbdf73607f240	1905	0f692345829560ccc163	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6f927d7c60ca1e7aaa1e	1905	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1616faf8eb64900579a3	1905	e48a2537426fd384e04b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ce94c56e6bda3a50c83f	1905	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e9830619628c76ea865b	1905	0a30a4d85b028245ef62	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
45564d9566a031f3c906	1905	028f24b1257fd6f8b5af	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ccb696eaa3bbb2ad60dc	1905	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
275e1ee11561e36a96d7	2403	13cf5f07e77de99dcc30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a2bab05e9513b9683c2d	2403	cf5177242c1034802d8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4fa12bfa181d402693d6	2403	ec964e12ec298f154906	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
722a0b2c4d3162c83cf3	2403	7f18686d985523c66aa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dfd01cb6a6070c994de5	2403	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
92fb81b4a4fa668684f8	2403	b232a97de45a99e40e65	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9d998b04b95481f67ddd	2403	50539e1ab2fe53f5df15	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c712e008bf25153511e1	2403	86c46654bedeb5a02591	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dfba1f8df175a9f60e43	2403	0f692345829560ccc163	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
78411db807f0cc9542d2	2403	faf6df81a0493135c916	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9e3f28ce5f32aa62e8a5	2403	e48a2537426fd384e04b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
14df638e2c638491d86f	2403	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
13f5616c88384946f4c2	2403	0a30a4d85b028245ef62	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0480d5accb2e350695c2	2403	028f24b1257fd6f8b5af	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4dcbf46f80f17d483961	2403	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6be2072ccd245526c5d1	84	8427cb0e609c6395011d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
79015cfda4f3339ee78c	84	cf5177242c1034802d8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
002255d43f51341ccfd1	84	ec964e12ec298f154906	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
28528c1581c75ad3a85e	84	7f18686d985523c66aa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b08f2f03cfcc980ed7b0	84	f6eeb477d5f1c37726c6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
64a26140672f249593ba	84	b232a97de45a99e40e65	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ce099314e1baf9e2fe93	84	50539e1ab2fe53f5df15	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
086dc89c90d534afb480	84	86c46654bedeb5a02591	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0de3ccf017dd0afd884b	84	0f692345829560ccc163	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b179a916100b57bd5d67	84	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f5979962c6390a806e8b	84	e48a2537426fd384e04b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1f85c86a245069d5b7dd	84	3ce1e9bfc95928b110dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4862d2d0f3e304ebf9cd	84	0a30a4d85b028245ef62	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
727218659f19461dd220	84	028f24b1257fd6f8b5af	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
270c39d56b7fb569c1f5	84	b96a13884d9cb2f18ad3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
344e5f41bf9d6f709492	2575	8427cb0e609c6395011d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
138abfcd39a345128adf	2575	cf5177242c1034802d8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fdcd36993e27cf9f63a3	2575	ec964e12ec298f154906	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e1134144bb78734ff027	2575	7f18686d985523c66aa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
427627fd7fa36f1154e5	2575	f6eeb477d5f1c37726c6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f44d8200444527e9c9d2	2575	b232a97de45a99e40e65	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cfd447000fb3feca31db	2575	50539e1ab2fe53f5df15	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
78ea4d07547c48364c19	2575	86c46654bedeb5a02591	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
64f40f5308ef5ffd27d2	2575	0f692345829560ccc163	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1b6165a1b0bb282aaffb	2575	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cc7efa75f99d17699563	2575	e48a2537426fd384e04b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1f86d7d6dea16c7a940e	2575	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
21741158696cc6053fc8	2575	0a30a4d85b028245ef62	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a24337cbd88468e4d0c2	2575	028f24b1257fd6f8b5af	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
390f45dfe705c01b9ea3	2575	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0129c67f572745a595bc	93	13cf5f07e77de99dcc30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
52a5c80944a9a812972e	93	cf5177242c1034802d8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b38aad46d38dff0fc392	93	ec964e12ec298f154906	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dabc5d5e77c5cd4731fb	93	7f18686d985523c66aa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
20d37f91e6e56181dea7	93	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9ed0e1db7c702e5949f2	93	b232a97de45a99e40e65	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
60ee55d9934882b08437	93	50539e1ab2fe53f5df15	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4221ac00fd8be20293f1	93	86c46654bedeb5a02591	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c06b27d8486f010c31b8	93	0f692345829560ccc163	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
70d6a34e45a4102a9b20	93	faf6df81a0493135c916	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
34ca22261e35fd2cff5f	93	e48a2537426fd384e04b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8d59bc29a5cfeaaff08b	93	3ce1e9bfc95928b110dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b13005ddb5243abb310c	93	0a30a4d85b028245ef62	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f99d0b9e4f3a17169583	93	028f24b1257fd6f8b5af	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d6a7362b455abe05948c	93	b96a13884d9cb2f18ad3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
912e4a9c6665e277663d	2219	8427cb0e609c6395011d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ad01464cf4e8b040ca69	2219	cf5177242c1034802d8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
876aa3b86a359343e124	2219	ec964e12ec298f154906	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f0d6ce58079369f0bab9	2219	7f18686d985523c66aa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b8185d5c358e9c98e43e	2219	f6eeb477d5f1c37726c6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
141e3bc8bff53453a7a9	2219	b232a97de45a99e40e65	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
30768f1e87f50caf6c82	2219	50539e1ab2fe53f5df15	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
30bad3842b1d6c15f2a3	2219	86c46654bedeb5a02591	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
30981804e202a0e3234a	2219	0f692345829560ccc163	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3f32556d75983ccd0245	2219	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
973a34cdc58c8bc44501	2219	e48a2537426fd384e04b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
caa55e9a1271611d043d	2219	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f6171c0163c211128030	2219	0a30a4d85b028245ef62	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a807fbe5ddf721c82d09	2219	028f24b1257fd6f8b5af	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c1cf31c1d359e04b8f77	2219	b96a13884d9cb2f18ad3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7bf705618e172b5f2445	782	13cf5f07e77de99dcc30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2cc3346e0efd736c7870	782	cf5177242c1034802d8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1f3e3bea1ff193519b57	782	ec964e12ec298f154906	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a44376f1a582d32a2a5e	782	7f18686d985523c66aa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f15d55a395e97c340ebb	782	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
599fdf55ad37457edb27	782	b232a97de45a99e40e65	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
59e89aa4b0152811bae6	782	50539e1ab2fe53f5df15	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
52e42727efa5a6813d79	782	86c46654bedeb5a02591	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
63185618c0efbf3efcfa	782	0f692345829560ccc163	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
48d768b44fae6428b606	782	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
635fcc916ec15684ffb6	782	e48a2537426fd384e04b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
76b40130d47b4aa2c4b9	782	3ce1e9bfc95928b110dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
24a746d54462a403a4d5	782	0a30a4d85b028245ef62	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9f6f53ef6f517a34e810	782	028f24b1257fd6f8b5af	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
675972d1a43080812bea	782	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
410b08ca85586c5756aa	1539	13cf5f07e77de99dcc30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e505b766fb5c30b94fb0	1539	cf5177242c1034802d8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
20fc785eefd01b1de6ea	1539	ec964e12ec298f154906	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0b574e7f7a241a0509f4	1539	7f18686d985523c66aa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
120d3053b0f0e13ff9ea	1539	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
26dc19a0035eb8574fdf	1539	b232a97de45a99e40e65	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5b1fd6ab7719cdb54e70	1539	50539e1ab2fe53f5df15	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c1c244bb5982cc09b97f	1539	86c46654bedeb5a02591	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3cb0cea771ca53b2b83f	1539	0f692345829560ccc163	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2c14467158c8bce2b986	1539	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e1e98cc3702fdafb0540	1539	e48a2537426fd384e04b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c6ea73afc5d839d7880a	1539	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b28e1927f3dedec904b6	1539	0a30a4d85b028245ef62	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aa3e9c9207919fa89d97	1539	028f24b1257fd6f8b5af	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
01b38e8b3b35c3dc578e	1539	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f5a4427b298ac9df6f5a	2579	13cf5f07e77de99dcc30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
386749f9472dfae7a828	2579	cf5177242c1034802d8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4289487abb75497a59f7	2579	ec964e12ec298f154906	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a0607920f927b8ddb956	2579	7f18686d985523c66aa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
030cb4dd1f9a8c9bbe62	2579	f6eeb477d5f1c37726c6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
179a4fd13bbb3b76224d	2579	b232a97de45a99e40e65	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0ee5de3ed0e063f60ce6	2579	50539e1ab2fe53f5df15	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fd0b84e249cb8eb067d8	2579	86c46654bedeb5a02591	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
71181e637cadad8ce3e4	2579	0f692345829560ccc163	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2cf8fc77d81bc0764bcd	2579	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
77d4a0e5ec430ac37dad	2579	e48a2537426fd384e04b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bde0527ce4c895787460	2579	3ce1e9bfc95928b110dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e7fa73703c04647c6170	2579	0a30a4d85b028245ef62	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
868e8a83fde4c1830941	2579	028f24b1257fd6f8b5af	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e28aa99a4b611d78cd03	2579	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f6e95dcad25d7dcc49d0	2347	13cf5f07e77de99dcc30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a510b41260b84fefc689	2347	cf5177242c1034802d8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
084857f7a528e694a8e5	2347	ec964e12ec298f154906	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e41b792606cec053755e	2347	7f18686d985523c66aa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8f02037588fab58b3ada	2347	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c5a288051d75c5951b5a	2347	b232a97de45a99e40e65	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9faf72d8b8b4ef709a6f	2347	50539e1ab2fe53f5df15	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cb421ee3fd832a42eabc	2347	86c46654bedeb5a02591	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ce722b57c5e19677c47a	2347	0f692345829560ccc163	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ef1091c3602192e99b88	2347	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
090b18b27dea4fec2ff4	2347	e48a2537426fd384e04b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0823fe22d16dc9d7a363	2347	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
439b29b461889cc6fbb0	2347	0a30a4d85b028245ef62	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bf28bd6763310ebdbb3f	2347	028f24b1257fd6f8b5af	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6be86dfc3d07884c29a7	2347	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5061dfdae4c6f78477ec	2146	13cf5f07e77de99dcc30	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c987987cc4ea9184eeb0	2146	cf5177242c1034802d8e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
02069a9d4f49f749cee3	2146	ec964e12ec298f154906	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
da72a8976de50f682927	2146	7f18686d985523c66aa3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
46d58fa6471a7614c29c	2146	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d9e8f1917038ae1db210	2146	b232a97de45a99e40e65	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
98d99b88648200e90f1b	2146	50539e1ab2fe53f5df15	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
781d784f38170b7d949f	2146	86c46654bedeb5a02591	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
db0eb6e5467466abcd9f	2146	0f692345829560ccc163	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cda6dccaa2f27a9299c4	2146	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6806404e3324d4d30b39	2146	e48a2537426fd384e04b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
64afbcfe12161d1cb11c	2146	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
36c9ec8cb0e5c6a3767d	2146	0a30a4d85b028245ef62	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e21ecc9b03e4531c50b0	2146	028f24b1257fd6f8b5af	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4379888b918df2830ee2	2146	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eab32c988e8e3f068d44	2029	809db13c3bb696a97ee7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
724cefb605d50dc6bc50	2029	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
398cb6a571a331ba239a	2029	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b848110ba51ef7f49940	2029	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4314e578d93db302571b	2029	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b241ff959b2a65e1e8b7	2029	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3508f9fffc2ba91a562e	2029	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f087ea697fd4364bc3aa	2029	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
09dd7bff87c15d7a5479	2029	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
14f1fe18aff8136320d7	2029	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a492ea4c71326dbb8234	2029	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
30e7e7b3e668da6d5343	2029	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
adb9a3b199f1a2c19505	2029	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
42f44d2f9d9b980aee60	2029	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
431f89e2d4ce331eb485	2029	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5fd077cd2b946930a4e6	1708	4d780077a193e1d55df4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
818884a9311abc2f1ec6	1708	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2fa590d4cbbb25c1a4c4	1708	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
67c22e8c1ed7eede8c8f	1708	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
742bae1f30c99a441a3b	1708	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1eb713e1afeda3dae787	1708	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4f0fff4baadd76ad2e11	1708	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
26be1b511f7c1c252940	1708	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
579909f57184cea468fa	1708	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ade2da69086e4f97b086	1708	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2e9f49a9f2c308aa3126	1708	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9b6cef8ff088cb32a2cb	1708	3ce1e9bfc95928b110dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0a7a5f351cc456b9ff20	1708	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ea841c6f4c0946b34557	1708	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
af57bae4f1cabcaf7472	1708	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cc0617d782e685f7cd25	2142	809db13c3bb696a97ee7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c04208240fecee1a036c	2142	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8334edc24a5af7b270ef	2142	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
71570d09b5e2f64b589b	2142	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a0b57af9aeed798444ca	2142	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
32d64cd35a6fb593ea77	2142	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d89d1f5959c577ef558b	2142	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
89e821d898784d16b967	2142	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a82b7a2486d9d7976979	2142	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c5f580c292bff24d1ba1	2142	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
00d9d8842c9a5e5719fd	2142	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
03621342ee2e3d1badb9	2142	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
af99337e348ccca0fc62	2142	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
25effea2a39aad0314c7	2142	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1b93204cc4f8ba5bea04	2142	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3c33a52a7a2badeb7ce1	959	4d780077a193e1d55df4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e5cbced011ebe05ec471	959	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
891222daf4cbf9c9b219	959	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ceaa274f7ede6c60cd72	959	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
42e2146fd77e3e1c37b1	959	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
685e421fd2aa0cb44d03	959	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3e0006afab15b08ab6db	959	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4762276f1ab4018d4244	959	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4ea2bfe98101818d3cac	959	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
da9c6f8f3f033bcd466a	959	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fc55990a9e7a96f64ac6	959	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d55cab9a4d9a2f517fa6	959	3ce1e9bfc95928b110dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aa08af69de9082be1c2c	959	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e6c1449e3b81e952fc28	959	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ce812b3c83b36e6bc825	959	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cc0ba15727e492ae00b1	1811	809db13c3bb696a97ee7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b78a29c87e20f5c4d6df	1811	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
289ad06bb789bbdd20fc	1811	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e352696ff8fe37c9a84a	1811	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c7ab77e3bb1e8a3a9166	1811	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
43a683d50f369d738cf9	1811	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8449068dc5459dc38502	1811	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
80f9616c963ecf7fec77	1811	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aead790745d1e024ae90	1811	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d078e51c31df55fe703f	1811	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1e8c665f4ee5154be614	1811	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4360979b64e7915efbf1	1811	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3c1a3f814b940ca27ae8	1811	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f7ea947aa39e9946873f	1811	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee32f1c29b68ca93826e	1811	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6a5540008202dc4284d8	2226	809db13c3bb696a97ee7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
57d8b1f7b17a9edbe4a1	2226	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9563159f43c0d0a89655	2226	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f0be56c06f601d5b5543	2226	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
56e6b0405e5be555c89b	2226	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
823573f5f3fc2bdd255b	2226	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ece6f77f714a30fe6a1c	2226	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e702ba0428706ab5ae93	2226	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1b97babcc531ec146f2f	2226	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1bfdf2942e35b79778ed	2226	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ba00d96467ac0fccc35e	2226	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3d84776202fc02a9b559	2226	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
62a6b7183cbe079d436e	2226	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ed402a290cb6b962b6bb	2226	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7cfb73776be5f4699cc5	2226	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
215a566c17ff6d9f2a11	74	809db13c3bb696a97ee7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
74e879c961d6e287ec9b	74	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
309edc5a1a6b3df8fcd9	74	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
17cd18eb296151a02dce	74	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
efa55f81acb987f152f8	74	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
302dd61ddad1a6c4f3c7	74	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ab130d004436fd1ade35	74	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b58f034d6f10f80e2e04	74	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b747d79f578a8f6793ba	74	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5009475947bf66b280e4	74	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4421230ff9f294d365f1	74	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cbe9ce38ef3475935188	74	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
832647a69b843c8e0803	74	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
87ffe40efd3de950df22	74	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7e9a54c3e71c256240ee	74	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
674ccfe67df31a3d30e2	1188	4d780077a193e1d55df4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
07300f5e28b3fe29c1c4	1188	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cc09e682146d5bc08e28	1188	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a31e40ef6140e90fced4	1188	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
60ba7701b214927714f8	1188	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6e14d07111326d6b31c7	1188	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
07c5a2ad2f04f8638d29	1188	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1ddc1ab6c2c1ebb726db	1188	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0494753fed247e2cb8e8	1188	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e815688369e1b9226d86	1188	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
45d97781344f3a5f1488	1188	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0167881a6fe4a163ba36	1188	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2f557678631aa9af9515	1188	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1ee4f79460d1cf06a51a	1188	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
59d2b061dacc38fd769c	1188	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
963f2b3172bd9d927ff2	1839	809db13c3bb696a97ee7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f8b34e74f5eb67bf01f6	1839	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
45e6ca05a3a0c7da365e	1839	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0accc2186c6e0f73e71f	1839	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fabb9d6719062ce20a2b	1839	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9cf388dae91981fc7298	1839	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d566964fa028a0a62c3b	1839	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
05cf8c4f6267fc6b2c4a	1839	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3f998752fc5650e4db85	1839	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
325a5709aac34fc62380	1839	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2994415c8d931e5f8fb1	1839	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
88cd850510022b8e696f	1839	3ce1e9bfc95928b110dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
085de1a61b2586a11337	1839	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9dc087ffcb232ea82812	1839	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
96cf7ba055426920e566	1839	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0c74accffd2457a99baa	1715	4d780077a193e1d55df4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7efb8c908e4f8a8f6736	1715	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d4c3680d80d2d9d16e13	1715	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aab327fe1bd78f6e79fc	1715	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ad6f75b3c2ceff14e944	1715	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c12cb61f9ec5360f4f40	1715	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0653e78031337dc5405a	1715	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4ad50b82d448a12f3c4f	1715	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8b97a36a20ca74aa3e9f	1715	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
209398286efa1bcf597f	1715	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7d025f3a193e749d4782	1715	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2603af8fc3c8c904feea	1715	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2ca9c7c915ba4f44655d	1715	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
92c1b2d69a130ba7cbcd	1715	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
686beda8b01bae6a3493	1715	2bd14e1ccaaa323968fe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
56d615ae9e01076a58e7	1189	809db13c3bb696a97ee7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
33f1e2fcacf8cc4ee7cc	1189	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7619823cfde8383e900a	1189	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b8ad49e3d6a80d5dea1a	1189	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bff229de55c6950a996a	1189	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a7e42114cd8407bf5ad4	1189	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
985a6bb7140895205a32	1189	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
953303666098e4d1561a	1189	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
80e969974d6e9fd56434	1189	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8ecadea3d918b166b3e8	1189	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2e19ae925381c541b518	1189	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ba1eaf04d2924fd4cf85	1189	3ce1e9bfc95928b110dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
50838f319bba46ecff17	1189	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e7a11f5cfce0b52ef1ad	1189	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
083b3bb71cf99e233e60	1189	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
affe686e666c004ef62a	94	809db13c3bb696a97ee7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4e880833cfa12a71518f	94	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
94f46c1889fe47ab79b1	94	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
854bee01a6ccce9e485a	94	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
987293ea5cd0b7038f1c	94	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7315511a2200c5844123	94	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c64a17225a473042bae4	94	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
af2dc267142421890dff	94	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
24351987225d5a94f80f	94	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
36007be8c035fbc132b0	94	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
681470e602cdb6109490	94	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6c875e293502ffe3260d	94	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a952e4d489b13e83e22e	94	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
01ea0ea13065a3e7cf49	94	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bd5ffacf7d5d36b10f62	94	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
664411bfad30d6187e85	97	4d780077a193e1d55df4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
82befa9269e271637333	97	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b1fcf221bde068a08564	97	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7a730b0da8c9111dc949	97	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1ed1bf9ad834b2ca6f8c	97	f6eeb477d5f1c37726c6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9a04eecbf4fc09c42b54	97	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5df5b413c96d74e2e824	97	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a188264b469194d50482	97	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
84e7e69494fd62b156b4	97	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
52c84808730ebeb9ce2b	97	faf6df81a0493135c916	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b1b8d3184fe4079d421b	97	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
957f2d32dc7df930a97a	97	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f05cffe7d67e02f4cd92	97	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dafae0d25103b3dc2440	97	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
625076084edf02aa0d32	97	b96a13884d9cb2f18ad3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7775062778725e499aa7	1193	4d780077a193e1d55df4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
898e00cdd892d3287afc	1193	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
be522ec8168cef6a2ec6	1193	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b5192a0c4da36ded4ff5	1193	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9f10147db32e4dc8e541	1193	9bb832cf61494f6878ed	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
58df721a95142819740c	1193	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0376dd32d5f4bc5a80b3	1193	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eb32fcb82bac93f17551	1193	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d508d5adad96746be197	1193	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
77ef2f2872047ec77ff2	1193	b95cadce5af0089d8b98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
afef028f466cfbef650d	1193	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d5ba0349d185b979861e	1193	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9d863870fa4795a4523a	1193	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c2a19bcd081706a9fdea	1193	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ac068a3239cd3d589e0d	1193	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
44969a0a4bf20db56db0	102	4d780077a193e1d55df4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5609390f446e0f2011cb	102	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
72007c6658eb541f7cba	102	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2ce62d613a6a39a288ac	102	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e9d7e194f2e9d847034e	102	f6eeb477d5f1c37726c6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
02be7e7a86c4a60f68f7	102	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8b32f90b2d59b4447484	102	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
851649d906614c16491a	102	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0ce62110ee751f9c4cdd	102	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d88260d1c94b1ff65156	102	b07f3d9674f357d4fd0d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4252c9d2d90de928224c	102	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3091aac64c94e6211e3a	102	3ce1e9bfc95928b110dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2ef9e6c54d176e7be709	102	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0e4cc301136c041ba149	102	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e16270cc929d9243c33c	102	b96a13884d9cb2f18ad3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0ae977ae183a659a25c1	969	809db13c3bb696a97ee7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6e954bf7d6d3273d3c1a	969	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
10c6e6278171897e5c8a	969	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dd2dbb4a6841527c2d98	969	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8678daa35c8859b3988f	969	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d7f6b615c9e77afaa5bb	969	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3f6febaf6604286e59ce	969	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
03072a874fd5ac1508e8	969	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e8d6c5f4295156f5616d	969	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d3ef21fd310b6717cb83	969	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
26bf026c17553e0cdcea	969	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
40f160f60117834f8732	969	3ce1e9bfc95928b110dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ca3d697b5169e71b5448	969	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1447639b22b4c42bff4e	969	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
797371d90a30b517baa7	969	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f4a2f729f5c6645dfd61	1133	809db13c3bb696a97ee7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
691d4ab4e976282c7f4c	1133	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0827a099a632c6f69f8b	1133	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
121b2fdd353ebf8e9643	1133	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
90febc72a9dfcce1058d	1133	5c960d1023e68e8b8d68	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2b515aa9f4d456b518df	1133	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
15a05c90f9d58178dcd6	1133	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
98f216bcca2cc61719a7	1133	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d8238e9a3b426f8ac498	1133	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
26c85dc8af835196a8ee	1133	faf6df81a0493135c916	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
07f528910913a833f2f4	1133	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
88f6b0b56911a181e61f	1133	c5aa80e1a22c4b09fbfe	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d2fe0f81354b23717715	1133	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f9f5f215fb416dd5e7c1	1133	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
af691547069dd0b11845	1133	b96a13884d9cb2f18ad3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
52d77264979a4931b9eb	110	4d780077a193e1d55df4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
552960c423fa292c350d	110	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6380411184cb3cab34d5	110	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eee971f1bea2b6a3325b	110	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
568db4dc06dcfdf23daf	110	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7e5a4d60b565ac6ff3c5	110	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6da822574690b62a50af	110	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9f6a21fcab0d2e8e915f	110	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
af4260629467c2b5b6a2	110	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e8ae226670a20274f007	110	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6195073cfae7cbb56a82	110	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
524c0d18e367a6904d25	110	3ce1e9bfc95928b110dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b7725790bc0e4d7760ec	110	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d95bb814a9239aeebf3b	110	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fd0b5f560cd825f30758	110	3d2845b4e39b80d24680	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6e0f925ad6898aded3a7	1830	809db13c3bb696a97ee7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5c22c8c947720a9ff65c	1830	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
175939028a760635ce7b	1830	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7970aad3456c5ebef594	1830	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cff81942ad2c27ce8181	1830	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
25bef9cf8f60827b62d7	1830	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
56da002c4869f3d514b3	1830	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a45652900cdcab676c4d	1830	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9156ec6ce70b6bf2c5bc	1830	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cd710158145ed63bc31c	1830	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5a35d48375da9f58ac01	1830	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c87d678f7008dd3e711d	1830	3ce1e9bfc95928b110dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a75b2791708fb037cef5	1830	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
de319325feed9ab418c7	1830	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dba2d741977745e6541a	1830	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8a3237fe7b30a740b345	1592	4d780077a193e1d55df4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3ea9e9cb8000f902147b	1592	1094e4c9fe718bcb6290	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1d3c0ad295e072249a58	1592	65d085a17aba3a9b1d94	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
748bce735a3ef930d276	1592	ab845c01757acc0d395d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
285eae8bdfccad8333b6	1592	91df72a6767bb7c3ab87	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c3d9108acd6d66601574	1592	b2e5e7431c984ac1c0d7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a5739f43d61be7d7b8de	1592	81f20f3be795fb6d955d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2999829c2fd45bc6b10c	1592	9e72e29c3c0b838bfc4c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8625095847491296d681	1592	19cc8e8396329ca6c0bc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
99d24d0faac64494cd99	1592	3c98fafabb219e9ecd4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e5772506cb40cbb20ed8	1592	5d13ab8d40255bc00b5e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
45a99c971537ea7662a7	1592	3ce1e9bfc95928b110dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
848d97d965908190000a	1592	4e4f54ae37e0e224fdf3	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3cd9183a0613e7930333	1592	acb5f26d24423de54ac7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2e93e1b2c3486d40f5f7	1592	67b3b7100d976861d8b7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
28c2b7338c7f9ea96ff4	1542	528f7d8434d6f3a6f7ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
99b8a74bf374cb717fa0	1542	283a8218e792eb7044a2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
64273dc3de088f86ea9e	1542	fc3cc49d60e02ea77ee2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
55fece6f2abed25bd727	1542	848c20d72e2e199b8639	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a46321b128bf29d2cd66	1542	6d7d4c292566a7ad40ef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a93e6898128eb1b57a45	1542	3e4a03a8fd0bb2afc961	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
36913f6082444308c60c	1542	ebee19666521dd14842b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0d37514368af7b486c67	1542	ec74201ba1740e7d0f91	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d72a87d983aa0177163f	1542	2e959c074f854901899e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5097d4f174ec56fc87c5	1542	a0eea0aac26bc2c9a33b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a6356a611f7c8719f92e	2304	848c20d72e2e199b8639	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fb7fbefb8f69b1b663e9	2304	3f5084d55890068aef0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6b66b2f5b65dea377f0c	2304	159a5eac4538c61ab219	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2cb47d19264a3852d4be	2304	69a272fe2ae16f6f7a3d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
655ba677e9be565f9e39	2304	d0011180a38970fb74e7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0e805a75af54d042a58a	2304	62216e3c0b97d44041eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
748e6495b44bd8eef343	2304	c5a5a84cddf882001a60	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
98e6663f7ebb68c7e5f5	2304	24076f2e2af052db9360	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ce21c3697d93d6a0b323	2304	fd0e3630e13db624ea01	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0df1539005547c0e650e	2304	a0eea0aac26bc2c9a33b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
239e1ff898a2a8447150	2335	5315b384ce058e8ab8a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
21dae62bbc5402643f28	2335	8082e012c8ef47aeca2d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f66599ae58725a7f09e3	2335	bd14ad4dd640660f1536	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
660b456115bd0bbd5568	2335	0e647084d6478b6ffe86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6915a7620a5733416814	2335	848c20d72e2e199b8639	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
13019b55404c948126e8	2335	6d7d4c292566a7ad40ef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3fcb94f78d880a62cd0d	2335	3e4a03a8fd0bb2afc961	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a904c57c1d14cea4e202	2335	ec74201ba1740e7d0f91	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f3fb81408b78e66d8c44	2335	c26810bb7ffda87028cd	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d001337f0ef2c791ac8b	2335	a0eea0aac26bc2c9a33b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3df20749324b06d430b9	2178	528f7d8434d6f3a6f7ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
008ebc43c6942c9d1b43	2178	848c20d72e2e199b8639	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3c7f4f68ae538ea59be3	2178	c44e345dd4e716d24329	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
40c2d784819142c3f9f3	2178	fda056455d00bf047c21	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f5396ac491f2ba56c4e8	2178	69a272fe2ae16f6f7a3d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dc1a9b7a023d8c2bf93d	2178	62216e3c0b97d44041eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
79a19a5b832d50e0a31d	2178	dd1b0f7f6f3aefb5955c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3bc82b5e17f88729e341	2178	24076f2e2af052db9360	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
391bfbd19f5013046e3d	2178	fd0e3630e13db624ea01	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b9c27312da0b0493b30a	2178	a0eea0aac26bc2c9a33b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4c9d2b7fc21ca99c301f	1949	5315b384ce058e8ab8a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
37739204870825acb4a4	1949	8082e012c8ef47aeca2d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4fa553890f538cad7030	1949	bd14ad4dd640660f1536	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0a0c5f1a1e42fa89cace	1949	0e647084d6478b6ffe86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
be64bfc0b5e5f756c3fa	1949	848c20d72e2e199b8639	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d56a0f16856927773cbc	1949	24f8f79de7d8daf5cdb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e2e38c8f30139ffd1be6	1949	3e4a03a8fd0bb2afc961	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
793d8618dd7377d74c63	1949	ec74201ba1740e7d0f91	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ebef8c68eaeeb752bf08	1949	c26810bb7ffda87028cd	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bfb5779dc797c792f232	1949	a0eea0aac26bc2c9a33b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
48ce33f8bc27b5eb50d0	2298	e48d91320929058cd529	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
52e9a539fa48fa2821fa	2298	848c20d72e2e199b8639	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9b344fd308b82fd75562	2298	c44e345dd4e716d24329	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
92b9372aaed27ca31402	2298	fda056455d00bf047c21	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8a429b2233cff6a4532e	2298	4b02a403546e54cc7f5b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e78ee9939d09179b2183	2298	d6bb5024d5ec592d3c44	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8b56553f4e1274f396bf	2298	d0011180a38970fb74e7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cdc5081d5717fbffb76c	2298	d7e3398ffa52ca717c5c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
abf382e6ee1c9ceccb4d	2298	24076f2e2af052db9360	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
816d3d239320ad920a50	2298	5f5846cdcdc67e6b6535	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
059ef6f4a471b8ba6023	1087	e48d91320929058cd529	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
df5057c8f9fe32c0c0a3	1087	848c20d72e2e199b8639	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e87bcf23bf7beb806a15	1087	3f5084d55890068aef0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fe69f85d77a0c070de01	1087	159a5eac4538c61ab219	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bb8f8e01c3133148a8e6	1087	6d7d4c292566a7ad40ef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7b1a22338d623d5d53e1	1087	62216e3c0b97d44041eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
15b367eacedb02dbc502	1087	3e4a03a8fd0bb2afc961	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5e53cdc2d65180d6fa47	1087	24076f2e2af052db9360	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
37e9a9b0a568bc001c4f	1087	b4234235efeea20a0ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1e93df6627d78a941b7b	1087	5f5846cdcdc67e6b6535	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b914870c0ecb1996048f	129	e48d91320929058cd529	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f03ce3fe5941a91f2308	129	848c20d72e2e199b8639	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c7afd0809f17039a5d8c	129	c44e345dd4e716d24329	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0177cbbe7af9911934ca	129	fda056455d00bf047c21	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c504fb94923a4f2561f0	129	4b02a403546e54cc7f5b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3c58f675759a2c915761	129	d0011180a38970fb74e7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1ab482f3724844f8d13f	129	62216e3c0b97d44041eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
af370b71ecc47684169b	129	d7e3398ffa52ca717c5c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7169f7447b7630c1d789	129	ec74201ba1740e7d0f91	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ae41313a94a22b1348c5	129	5f5846cdcdc67e6b6535	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
66cd358b42bfea9daca6	1929	528f7d8434d6f3a6f7ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1db317315c8b41e72cc4	1929	5315b384ce058e8ab8a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8aa0680ca957427f7c68	1929	bd14ad4dd640660f1536	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1d9fa4faf581bf676a9a	1929	848c20d72e2e199b8639	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f0e44cf7624a68ac7940	1929	6d7d4c292566a7ad40ef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d2c24073b42e1d17e7c1	1929	3e4a03a8fd0bb2afc961	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7e3b07da02e2ebfbf5b8	1929	ebee19666521dd14842b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8840bf7523671a20a0a1	1929	24076f2e2af052db9360	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c8680c7a48c85c307e30	1929	c26810bb7ffda87028cd	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
00426966e71479acf486	1929	5f5846cdcdc67e6b6535	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b3290479823dce2175cc	2030	528f7d8434d6f3a6f7ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6fff077cdd122f826630	2030	5315b384ce058e8ab8a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
699d4a74556ed96ee708	2030	bd14ad4dd640660f1536	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b7e920060b38735ae9ad	2030	0e647084d6478b6ffe86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ca81a55c115f411f99e6	2030	848c20d72e2e199b8639	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6c3927ad9b89c3df1456	2030	6d7d4c292566a7ad40ef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
83f10fb7f10d94aeaa7d	2030	3e4a03a8fd0bb2afc961	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
84067e5b55fb4c2feabf	2030	24076f2e2af052db9360	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
854e7e9977fb05d4228e	2030	c26810bb7ffda87028cd	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cd3c59e20bad5039fd7a	2030	5f5846cdcdc67e6b6535	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
182d7a7c8bf1402f5e6c	2343	283a8218e792eb7044a2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d59c721a16db07bb9b11	2343	8082e012c8ef47aeca2d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
25d56b8240d66fa0079c	2343	bd14ad4dd640660f1536	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
12967b448fd042a26f90	2343	0e647084d6478b6ffe86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1ea7e11c3781b41d432a	2343	848c20d72e2e199b8639	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
579b5874113a45fd29ac	2343	24f8f79de7d8daf5cdb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
02e89000e7ceaddea2df	2343	c5a5a84cddf882001a60	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee3663c10199ed3c4a87	2343	ec74201ba1740e7d0f91	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ae6bcaec0edddeeb9f67	2343	2e959c074f854901899e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
79202c8d2487dcec65d7	2343	5f5846cdcdc67e6b6535	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4d7cf384fb687df90d8f	130	5315b384ce058e8ab8a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8744f0dfb8a9a65ea172	130	bd14ad4dd640660f1536	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c0ac99594edcef1bc0da	130	848c20d72e2e199b8639	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c677a269e20dad89f8fd	130	6d7d4c292566a7ad40ef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
059e7450684d0218aaa2	130	d0011180a38970fb74e7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a82b7511aa93bf5ce389	130	c5a5a84cddf882001a60	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5b104916138ff9296c4d	130	ebee19666521dd14842b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ae738929ca02c49cef6b	130	ec74201ba1740e7d0f91	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a4ebd6c5e8b765f2b3e7	130	c26810bb7ffda87028cd	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee59398d01e0e156549d	130	81ceff273e6f7f288ad5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
db1c96f5bcf552b9e9b8	2152	528f7d8434d6f3a6f7ad	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8dfe06e504531e163704	2152	e48d91320929058cd529	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
183f8e51782abaa06a99	2152	848c20d72e2e199b8639	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7afdcb3402b691001228	2152	3f5084d55890068aef0f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
301d3569f57491ca4b47	2152	159a5eac4538c61ab219	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
725a28743077f3b7d451	2152	4b02a403546e54cc7f5b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e7eed49eba53d3f584b2	2152	d6bb5024d5ec592d3c44	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9f5c8c77062e72dd2313	2152	c5a5a84cddf882001a60	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f62435e8955311a0ed39	2152	24076f2e2af052db9360	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cfff096a01e14c01abee	2152	81ceff273e6f7f288ad5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
11375e43f5076b8a3e57	162	283a8218e792eb7044a2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ae8b02b8939e097d02e0	162	fc3cc49d60e02ea77ee2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
97dfdef2cb6abdc104a9	162	848c20d72e2e199b8639	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f8d4a3229414d9b16079	162	69a272fe2ae16f6f7a3d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7623bf595fedc89ddf08	162	51b76e207ca154e8acba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7125e2cdc8686c0ca1d7	162	c5a5a84cddf882001a60	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
189941ae0b8f911c412d	162	24076f2e2af052db9360	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aafb9295bac5027f73f3	162	2e959c074f854901899e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
69a397729ffced9c80a5	162	b4234235efeea20a0ca7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d40c6f331129e0ba0321	162	81ceff273e6f7f288ad5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
91d8051925fa87303ef8	1817	5315b384ce058e8ab8a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
98607576cff3b7d3e9f0	1817	8082e012c8ef47aeca2d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
694b6aaba57fc15597a8	1817	bd14ad4dd640660f1536	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b873de70db6d2b25d466	1817	0e647084d6478b6ffe86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f6aaada4db1e55482308	1817	848c20d72e2e199b8639	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bfa3ea15179b45ba028a	1817	24f8f79de7d8daf5cdb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5ba3ee98ae1ff73048c7	1817	3e4a03a8fd0bb2afc961	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ae5be7fb007610ea7a90	1817	ec74201ba1740e7d0f91	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
145075f27667a6f04c3e	1817	c26810bb7ffda87028cd	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
62fd9a6bf7bf9436e3b3	1817	81ceff273e6f7f288ad5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3b910512ba98bdc2bc49	188	283a8218e792eb7044a2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
440c6eb9d28f8aff7c14	188	fc3cc49d60e02ea77ee2	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d22f6fef4d510ca65682	188	848c20d72e2e199b8639	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9719fca9344a7c3184bb	188	24f8f79de7d8daf5cdb1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
00b42f17dbe67ffc4cb2	188	d0011180a38970fb74e7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f79eb3494d0b0619de30	188	51b76e207ca154e8acba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f3465d6a80a94c924728	188	c5a5a84cddf882001a60	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a90c3317d52567dd9f17	188	24076f2e2af052db9360	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e4b5e2b741d6039acb50	188	2e959c074f854901899e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ffe9cd64f60ef76da908	188	81ceff273e6f7f288ad5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8511091a8fa4af7e2132	2158	5315b384ce058e8ab8a6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
32c05927b269c320ac80	2158	8082e012c8ef47aeca2d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
155c9ce48154e856daed	2158	bd14ad4dd640660f1536	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b2936c78e093f940113b	2158	0e647084d6478b6ffe86	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a6444841357416a364c1	2158	848c20d72e2e199b8639	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5e7311b94344b321cc00	2158	6d7d4c292566a7ad40ef	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6e1d81b28b818c498d69	2158	3e4a03a8fd0bb2afc961	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6f0d13de7ceb0afa41c5	2158	ec74201ba1740e7d0f91	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
78b5df1993f12e738183	2158	c26810bb7ffda87028cd	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d2662b99f0c312576b42	2158	81ceff273e6f7f288ad5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d7b631d5c132bbd8c7b9	2224	dfb67ad547da1c74dec0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
10a4ae096281a8a52ea7	2224	366c6f405b1f8bc78439	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b32336b4e0dd62947b05	2224	08cda36fbe812c092c9c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b4d02250803cc6632901	2224	f27cf82c1214ada6fb4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dd03aa61e37d2645b000	2224	0bfa86e7a44f47e58019	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
280f3f4d192a63f0b874	2224	ffc06f66d5b1395c4fe6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d6124cebca9d462e1561	2224	ff8a7abaa5467496e924	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a155bfe89d2d03395968	2224	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
edd55a4be27c185ba6b0	2224	f212ea07533e401eb4df	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f9e7715dc9e84faaba31	2224	fbd0f00af5ce74d958eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
654fec8718f693a07f62	1020	53889c63251a67f47a2c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f2a7cea21b0672ebbd33	1020	4b963bec52944fc95745	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
230b5f58d7fd61666ae7	1020	af2892f8639e42fdc83d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3388cb259690166f8e9d	1020	f27cf82c1214ada6fb4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c2c2238cd0da0dc02fd1	1020	d131a3b33401a7b4a668	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
844bccb2b30385fbfc64	1020	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fa3ee7f5d47e1687b0eb	1020	9035d56968d798f556cb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d08ea34b7adf0e7694c4	1020	7788942930b1d0ff75cf	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ad59774f83b2b408372e	1020	591c7cad2de417c3536e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
594b605fc6d36ae007bb	1020	fbd0f00af5ce74d958eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a251b0dd084c559d9d5c	314	cbbb379c9906538aaaea	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2149fca1283ad1ee70d5	314	02938d7b62625b948141	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
15694088396dabbec57a	314	6189d54b165dab8bd56e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0f201f75a0178230bba9	314	a82276e194b1732711b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cad3ab211b3142803b86	314	ecf47a0bf5f29be17314	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
af51d41d062ab40ae691	314	15341f322b282b495094	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
026325ecf44f5ec440f0	314	ede2ca1d19344a1b8138	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7871171fe0a8f9452482	314	7788942930b1d0ff75cf	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7afdb26930586e44fab6	314	9ea43fbf389f497e0a95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8cc270ef76d5854b7222	314	fbd0f00af5ce74d958eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
18216b309ad9abc130d2	2214	cbbb379c9906538aaaea	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0e10afddc267cd152741	2214	02938d7b62625b948141	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b044cbf40dfc74c7eaae	2214	4b963bec52944fc95745	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
11754e5e7849e97f16e5	2214	af2892f8639e42fdc83d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1288971b06d7cc4ce114	2214	d131a3b33401a7b4a668	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2bc8b88e59ebd3d6b9e9	2214	15341f322b282b495094	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e8536a5826ba69b02c34	2214	9035d56968d798f556cb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
74bec68928face4efd6b	2214	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d8770ff362864d2df7fc	2214	9ea43fbf389f497e0a95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1c1293d014c4499f2ee5	2214	fbd0f00af5ce74d958eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7d66db7a0242197ba24a	318	98cc7263651f31776990	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7710a054192f15adbb46	318	366c6f405b1f8bc78439	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4d1627b944ead99f6980	318	8ce3f350dfe12ca44f95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
98ed4384d891b23af25e	318	02938d7b62625b948141	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d36198a86f2b3af43905	318	1b29574d86db3b63e11e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7cc33e61e7c691e6b124	318	15341f322b282b495094	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
75529c66e1290f081833	318	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fe34163ba504f2e4a7aa	318	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7f31824935db8aa98b17	318	8ecde88989703bbe399b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dd04b56a4e0685874165	318	fbd0f00af5ce74d958eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e5f95317eb6702bf0310	944	98cc7263651f31776990	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6d1c3cd678a6def881ec	944	366c6f405b1f8bc78439	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
44c34a1a1def928031f4	944	8ce3f350dfe12ca44f95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
89228f06e77bd7dd92a7	944	53889c63251a67f47a2c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0cd562682a1fd8373432	944	02938d7b62625b948141	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e79eb0e11afb4f8f2875	944	d131a3b33401a7b4a668	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d45834ea8d2c8b85234d	944	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
23b85d8d7dc007c6bfbd	944	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4708b3a0bc1981b32d8e	944	8ecde88989703bbe399b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
08ab34440e2a79417522	944	fbd0f00af5ce74d958eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
88a45433b8d389bb180a	2162	cbbb379c9906538aaaea	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7b19dfa079f853c48c94	2162	4b963bec52944fc95745	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6013557587ae38a1e0e1	2162	af2892f8639e42fdc83d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6dfaef775c5c873c3987	2162	ecf47a0bf5f29be17314	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f1144cab90951dcb3669	2162	e295ccde749fd2982eeb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
65d98f0791de0ae5a98b	2162	ffc06f66d5b1395c4fe6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
949c7e7e301b4fa08ef0	2162	ff8a7abaa5467496e924	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1a91d3979b1e05fc9c90	2162	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d87c0f4c4033712ea69f	2162	9ea43fbf389f497e0a95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c6fa4a9aee771d8fe7cd	2162	fbd0f00af5ce74d958eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c8558e1b84a635312010	1802	dfb67ad547da1c74dec0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
06e434564ec0dd9642b0	1802	08cda36fbe812c092c9c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5a562a51f969dc1097ba	1802	53889c63251a67f47a2c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2895d5c9e3fee40fa58e	1802	d131a3b33401a7b4a668	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9a2cd11bb610a4c85924	1802	e295ccde749fd2982eeb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6a59b276c16130e729a3	1802	9035d56968d798f556cb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e90630a31e069d3f1661	1802	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2e7a707de7ebefe1b2c4	1802	f212ea07533e401eb4df	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fae9072428a3c401b27c	1802	9ea43fbf389f497e0a95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aa685f3d53caeaafb152	1802	fbd0f00af5ce74d958eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
939a3de41db41e2103b6	1022	cbbb379c9906538aaaea	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fcb27e0087be58978ced	1022	4b963bec52944fc95745	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a8c076b161e088021242	1022	af2892f8639e42fdc83d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2b39cc513dd58b165f35	1022	f27cf82c1214ada6fb4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
46b71bddcbcff4e8d574	1022	1b29574d86db3b63e11e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
71591bb704167c582d2d	1022	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5336766c8bd65c9e5db8	1022	ffc06f66d5b1395c4fe6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e0bc58d8a7f47ed4419c	1022	9035d56968d798f556cb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f7a0e2ce7fa27028121a	1022	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c33d330a9bd4b1557278	1022	fbd0f00af5ce74d958eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b148a6ab66f9ad8c5851	970	dfb67ad547da1c74dec0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
da4147af23d0e0dc52a2	970	366c6f405b1f8bc78439	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8acbc69331b22dd12993	970	08cda36fbe812c092c9c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6251de955c413da551de	970	53889c63251a67f47a2c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0221973e14a633156244	970	02938d7b62625b948141	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
02748302a83bd14f2fc9	970	1b29574d86db3b63e11e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0f8e06adcf7812fa8e3e	970	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
29a520a15587e2a57040	970	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6901eb5873493ef61384	970	f212ea07533e401eb4df	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a3db86d719658fae599e	970	fbd0f00af5ce74d958eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
74beaf60546829326afd	2217	dfb67ad547da1c74dec0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e42139ba4d0bda16a492	2217	08cda36fbe812c092c9c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e563de21fcd11e67c120	2217	1b29574d86db3b63e11e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9c4a93d064f7ad30f170	2217	d13bbcf819ba7aedcab8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dbb8ee3331500c26f8c1	2217	15341f322b282b495094	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8ab090c2c8f149b5fca1	2217	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3f89b7746c4383ac4054	2217	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b261a486562fa5852d53	2217	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a70336e33c5fcee3b3ae	2217	f212ea07533e401eb4df	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c79326fd1c5c9ca11606	2217	fbd0f00af5ce74d958eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bf98a50c14b232695809	2365	c6854f08ccb4b2c8da4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9dfb6690344a644e840a	2365	6189d54b165dab8bd56e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5b2717a7b7b43bbfed1a	2365	a82276e194b1732711b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d70e8610ab8210ba0132	2365	1b29574d86db3b63e11e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ef8f6d6b6d6c5a356613	2365	e295ccde749fd2982eeb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d70bb2a75b1b0f8a656f	2365	ffc06f66d5b1395c4fe6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bc52f50944658ab8ac39	2365	9035d56968d798f556cb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8d9b4998aaad7d74b87d	2365	7788942930b1d0ff75cf	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aa5ecc9bdb2d530d60e2	2365	591c7cad2de417c3536e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8cdf47c4e7419ec0dc12	2365	fbd0f00af5ce74d958eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4cc37e981b35c5a946c1	1915	c6854f08ccb4b2c8da4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1ea1e4024275ff047f9b	1915	cbbb379c9906538aaaea	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
14d493934bc6454a951a	1915	6189d54b165dab8bd56e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0c7411b9cfe7d8f30b4e	1915	a82276e194b1732711b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9ec443213bd80887193d	1915	ecf47a0bf5f29be17314	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1e29473d233f2ac29f6f	1915	d13bbcf819ba7aedcab8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
55ec2cef6447f16e8ddd	1915	15341f322b282b495094	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
41392ff49920b91b4f05	1915	ede2ca1d19344a1b8138	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
838199685ed7a42d27fd	1915	7788942930b1d0ff75cf	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
156ae3cf6a0b895945f4	1915	fbd0f00af5ce74d958eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
802b5ae0c4c5b543c563	1999	dfb67ad547da1c74dec0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
64a04adff73602a4725c	1999	366c6f405b1f8bc78439	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
116fffbd7066e987043a	1999	08cda36fbe812c092c9c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1dc9d823f6a13ce5ed78	1999	02938d7b62625b948141	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
94fdbfaa9dabf1ded5d3	1999	1b29574d86db3b63e11e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2c838539a0a8ce93d4ce	1999	9035d56968d798f556cb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a2044d57bec47b72aa0a	1999	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2f6289e1d3721668b5d0	1999	f212ea07533e401eb4df	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
08355946c538931164b2	1999	fbd0f00af5ce74d958eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f3662ecf629d6a5cae12	1999	f3f87a4f6848eef065be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5d7743af5609afa0000e	1963	dfb67ad547da1c74dec0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9f7ae08be3c6fd2e4e7c	1963	366c6f405b1f8bc78439	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d328fbc7d32dda49e922	1963	08cda36fbe812c092c9c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bcafbe05a4516fddcb9c	1963	53889c63251a67f47a2c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fe64edde778175fbded0	1963	f27cf82c1214ada6fb4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9b9d597b379a7c8a712d	1963	ecf47a0bf5f29be17314	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dd4fb11d5b42898f8d11	1963	9035d56968d798f556cb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
079dd935666b49b97b1d	1963	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4c95d0b6b4643ae3c897	1963	f212ea07533e401eb4df	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2c7cdb1dc22672de2806	1963	fbd0f00af5ce74d958eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
896f96eaf80c9da1c300	343	dfb67ad547da1c74dec0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
19e517e74790bd5cc224	343	08cda36fbe812c092c9c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b688b5eb2119f867a23b	343	02938d7b62625b948141	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b859062fcad0e0d994f2	343	ecf47a0bf5f29be17314	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2bcd36d5eb1cdb0d4598	343	ffc06f66d5b1395c4fe6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
594798d7bec8d02ede01	343	9035d56968d798f556cb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bb019fefeead7d16c509	343	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a26a44d881697948e82e	343	8ecde88989703bbe399b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
deb65cecbc4cd2dabaa0	343	9ea43fbf389f497e0a95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a049346299bc72bd0579	343	fbd0f00af5ce74d958eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
513bf32cd52874c8418e	2181	dfb67ad547da1c74dec0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7c0eadd64b7ef98a3bc2	2181	8ce3f350dfe12ca44f95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bc60b8f165754cc7ba58	2181	d131a3b33401a7b4a668	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
98f3e85dc89a46bf1abb	2181	d13bbcf819ba7aedcab8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6efd8c1af0cba8c3caea	2181	15341f322b282b495094	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6509501b2e095c0c590d	2181	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f769a1bef7269f304961	2181	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3684b647539a40f5a0a7	2181	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1024b8d2ebfe25a34465	2181	f212ea07533e401eb4df	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
56a208404b3e07e7417f	2181	fbd0f00af5ce74d958eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1ced93fb7127e13fafd7	1919	c6854f08ccb4b2c8da4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9ec6c95fc1f0fc92dfd3	1919	dfb67ad547da1c74dec0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2939705e71232e45168a	1919	08cda36fbe812c092c9c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a165b67e8a31573f109d	1919	53889c63251a67f47a2c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5760b9795fa9910db355	1919	02938d7b62625b948141	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0d205d18f6c86a0d810f	1919	0bfa86e7a44f47e58019	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
17d07b0d14e3b36751e6	1919	ede2ca1d19344a1b8138	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0e1ca78d7e51da7026da	1919	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f9436a9af3337c69579d	1919	f212ea07533e401eb4df	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
335a001773d80c35a901	1919	fbd0f00af5ce74d958eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
794a438d5c46021ff6a7	1210	98cc7263651f31776990	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1dfa576d784936813f7a	1210	366c6f405b1f8bc78439	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a60f100016351f763a0e	1210	8ce3f350dfe12ca44f95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1e529dd1782742ae3e9c	1210	53889c63251a67f47a2c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fd479a32353382c353d9	1210	f27cf82c1214ada6fb4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
52a8b9e52fdab314256f	1210	1b29574d86db3b63e11e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c1a56c00bcdd5fa2234f	1210	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cf1ffe764c5313b84a3b	1210	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b0e498746f288e1ad1a8	1210	8ecde88989703bbe399b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4007ee1768a05af84e9b	1210	fbd0f00af5ce74d958eb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e4b9a39ffa5278e45e6a	1770	dfb67ad547da1c74dec0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4f53842606423b9a6078	1770	366c6f405b1f8bc78439	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
275e040e4858ac0f5659	1770	08cda36fbe812c092c9c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
adbd06e7cf765e1b57aa	1770	f27cf82c1214ada6fb4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
df82f3a23feadc472f07	1770	1b29574d86db3b63e11e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e86ffcb47eaffcacf43d	1770	ffc06f66d5b1395c4fe6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ed6aaf885e440107b75b	1770	9035d56968d798f556cb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
244a67cf128899f6b56e	1770	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
765b69193f820dc5f885	1770	f212ea07533e401eb4df	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
83d9bef461d6763230fe	1770	a89ce690562f1bfb0bba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5c1bd7da7f085a8e8e58	2160	dfb67ad547da1c74dec0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
de2f7503f050873cb287	2160	08cda36fbe812c092c9c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
903d6eb3fd3400a52bd1	2160	02938d7b62625b948141	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6c0479d78396d7f15c37	2160	1b29574d86db3b63e11e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bf1f26848e8641fe0ece	2160	9035d56968d798f556cb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
21e921312164ed40e96f	2160	7788942930b1d0ff75cf	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
97b63e62d3a652254357	2160	f212ea07533e401eb4df	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
64701c8898fc9f26480c	2160	9ea43fbf389f497e0a95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
eaffe9a89e61c3b5add4	2160	a89ce690562f1bfb0bba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
68b59b5ccd3385b4b0cd	2160	f3f87a4f6848eef065be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e08cb33f15d0eb263175	1726	98cc7263651f31776990	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9f4c1f2a855c1197d99c	1726	366c6f405b1f8bc78439	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e3e187c698ae384b225c	1726	8ce3f350dfe12ca44f95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
70c05d722a69f2f37474	1726	d131a3b33401a7b4a668	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0f6e754753f58ef78b5f	1726	e295ccde749fd2982eeb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2ec8be7b44ad26d71a62	1726	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a3c674c1c72cf163db77	1726	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3bc7a4aee4a5c362eaef	1726	8ecde88989703bbe399b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a2795572e47a7ee64b10	1726	a89ce690562f1bfb0bba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
58031eed8215033a4621	1726	f3f87a4f6848eef065be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6258450852595a638b87	1727	98cc7263651f31776990	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
be69f3c57ddef55a82a4	1727	8ce3f350dfe12ca44f95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7ed38ae16eda445db565	1727	53889c63251a67f47a2c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
112989ace62980981c99	1727	02938d7b62625b948141	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
057bde96b32b5e8e4e92	1727	d131a3b33401a7b4a668	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
adcfce63b363d2c5f2b2	1727	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c2765a3403db3f501258	1727	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4a3d2b3b3ad6f25c8eff	1727	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9179ef050d36637b9b1c	1727	8ecde88989703bbe399b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cd2baa267216b44cd544	1727	a89ce690562f1bfb0bba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e93b6dc9968b80f43a73	730	cbbb379c9906538aaaea	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0846f3f0cc660abbbbb9	730	4b963bec52944fc95745	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4fc2333ff8c0eaebff00	730	af2892f8639e42fdc83d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
de47749d931c884afb1c	730	1b29574d86db3b63e11e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9fc90e5d7fa888ee7b35	730	15341f322b282b495094	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aab4365e4340b648b633	730	e295ccde749fd2982eeb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cd4a377f9031d9146621	730	9035d56968d798f556cb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d0e8ad841dd5caddaf06	730	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6c0c170ddcc08f41dbe0	730	9ea43fbf389f497e0a95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ec682aa1da3d4f4e315c	730	a89ce690562f1bfb0bba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6235a63396ea1872faa1	2363	cbbb379c9906538aaaea	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
698163cdfe98714c99db	2363	6189d54b165dab8bd56e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8ad06e73e1ad699d2dcd	2363	a82276e194b1732711b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c85ae80291491ccfe0ce	2363	f27cf82c1214ada6fb4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
059f929f9f7d224b9a52	2363	ecf47a0bf5f29be17314	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0e9b5b6be764eb0456ee	2363	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b26b6817fefaffff3a1d	2363	ffc06f66d5b1395c4fe6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3aa9f395d3d0e05e2924	2363	ede2ca1d19344a1b8138	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c8a83e0d200f3970d0ca	2363	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4b067c2eccf697bab86a	2363	a89ce690562f1bfb0bba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
45875cb0812d13ad373e	2245	98cc7263651f31776990	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f58bd0900a9047d87074	2245	08cda36fbe812c092c9c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a0a398cd18de2712aa6c	2245	53889c63251a67f47a2c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
debaa8bcbf4d7ae9d6c6	2245	1b29574d86db3b63e11e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bb4a0649d03d13d95cce	2245	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8c5ee51451f7625040cd	2245	e295ccde749fd2982eeb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
71f0ff63240ecfa09e9d	2245	ff8a7abaa5467496e924	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a8cce1ec82fc4e1d5a27	2245	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6954aea3b4bdca305c97	2245	8ecde88989703bbe399b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
64e525a98223d6c67308	2245	a89ce690562f1bfb0bba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7d5df8d376e1aeec8ae2	319	cbbb379c9906538aaaea	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b345f30837b2d8ac9ae9	319	4b963bec52944fc95745	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
baacd651d9d9b789f5dc	319	af2892f8639e42fdc83d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
35e7295c06c60673b0d0	319	f27cf82c1214ada6fb4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3f25670128abd05ccbce	319	0bfa86e7a44f47e58019	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
89120920795764bedcb2	319	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f8fbecc9acdd861fb087	319	ffc06f66d5b1395c4fe6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
af7e0fb7928f813b3efc	319	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0a5777f629246162cecc	319	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f97aaebbb4d1ad8a156e	319	a89ce690562f1bfb0bba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fdafb4eea0dd04cf4d80	2161	dfb67ad547da1c74dec0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dfca10ad09fcb28e5705	2161	08cda36fbe812c092c9c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9679423592223096d6f3	2161	f27cf82c1214ada6fb4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
49a13be166b161bfbb25	2161	1b29574d86db3b63e11e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1f0a8ee4dc131025589d	2161	15341f322b282b495094	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f8e1e546a5db64952f7b	2161	9035d56968d798f556cb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ec1a845c271be57f45d6	2161	7788942930b1d0ff75cf	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5b0a3f72882b3436f721	2161	f212ea07533e401eb4df	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cf92b30aadd974d2d5ad	2161	9ea43fbf389f497e0a95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1ee5bb4eb42772186480	2161	a89ce690562f1bfb0bba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
efc5938793c8dd968f8a	1212	cbbb379c9906538aaaea	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
484d6f503389100a00f8	1212	02938d7b62625b948141	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6ed6b62347807ed7bf8c	1212	4b963bec52944fc95745	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1c72721fcc47edd74115	1212	af2892f8639e42fdc83d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
46812fe024a2739e32e9	1212	1b29574d86db3b63e11e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e044201b3153e94467d8	1212	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
09a16fc6d8ff35d3188f	1212	ffc06f66d5b1395c4fe6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8c6ac56f7c79c4051ec7	1212	9035d56968d798f556cb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bc345bc547fc381a0d2c	1212	7788942930b1d0ff75cf	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1262559c74f6f38c04c5	1212	a89ce690562f1bfb0bba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8d2aeccd116adc9bbf28	2239	98cc7263651f31776990	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
96abd5f19c9033f031d4	2239	8ce3f350dfe12ca44f95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
54b2dba5255337e4b82a	2239	d131a3b33401a7b4a668	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e407c2aa9ab1f6af4d52	2239	e295ccde749fd2982eeb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
398b7d5f380a296a38f0	2239	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e0fb9e5d4e7669e0c29f	2239	7788942930b1d0ff75cf	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
85ed23fad0a491dffb20	2239	8ecde88989703bbe399b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
293eef576dc681f55fa9	2239	9ea43fbf389f497e0a95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bb8d22184acd59f00828	2239	a89ce690562f1bfb0bba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8a825be76b31e8f0c7c4	2239	f3f87a4f6848eef065be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b4a61bf812a8a472fda1	329	cbbb379c9906538aaaea	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b2b020dcf309053bbd33	329	02938d7b62625b948141	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5d800bf8c63291de689b	329	6189d54b165dab8bd56e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
84a5cd973ab169f6f98a	329	a82276e194b1732711b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
77948bd2313482367dbf	329	0bfa86e7a44f47e58019	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
754b2059a1737e4475ad	329	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b1eb604d4da8ea9dcee9	329	ffc06f66d5b1395c4fe6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d17016c5da7c3e3182bb	329	9035d56968d798f556cb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
47aa3a4bc86bc35ed524	329	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c648422828e881982a3c	329	a89ce690562f1bfb0bba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
97a29b7232c61a409111	2164	cbbb379c9906538aaaea	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fc59dcff3a9b57e6166b	2164	4b963bec52944fc95745	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c74aafedd8a6773e5e7f	2164	af2892f8639e42fdc83d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d614fddfa3474fec76f0	2164	f27cf82c1214ada6fb4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c4f45c057d935ddec7a9	2164	ecf47a0bf5f29be17314	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
88feb98aa1a9453bcecb	2164	15341f322b282b495094	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8e65dacb99ad7fa6b64b	2164	9035d56968d798f556cb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dbea24809dc68806aa2a	2164	7788942930b1d0ff75cf	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
88547a82cc0b6ed1763a	2164	9ea43fbf389f497e0a95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
29b5b94f6def4fc5b8b7	2164	a89ce690562f1bfb0bba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
279751c625276d7bee76	1916	98cc7263651f31776990	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b0866fb5239873f4a379	1916	366c6f405b1f8bc78439	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
37bd7f348abcdbd52040	1916	8ce3f350dfe12ca44f95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
546d1d26b64ec8365a65	1916	53889c63251a67f47a2c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
05af22d7bfa87425c45e	1916	02938d7b62625b948141	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
70aba94e7dfd0ee0b348	1916	ecf47a0bf5f29be17314	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
58721d1562a86ff834ad	1916	9035d56968d798f556cb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dcc3d1424dd65f08184e	1916	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
765fe65a656866ce11c1	1916	f212ea07533e401eb4df	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5d927dca375a6df1e12c	1916	a89ce690562f1bfb0bba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
de86895354489f09a6d3	1917	dfb67ad547da1c74dec0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ce97b694a8ad579e2900	1917	08cda36fbe812c092c9c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4bbaa043452ca67445b2	1917	02938d7b62625b948141	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
07eab7131fa489b89b2d	1917	d131a3b33401a7b4a668	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
70c1d7358d865cce2855	1917	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3af499df8c1bbaf53dca	1917	ffc06f66d5b1395c4fe6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d49b93fdced9cf39e1ee	1917	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
746bb64be186be26ad8d	1917	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
77a30a555255e7d35d3e	1917	8ecde88989703bbe399b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b407c636526ec86b1916	1917	a89ce690562f1bfb0bba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3815e97785d17ecb3972	974	98cc7263651f31776990	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e24ad8f154f0d3a3f3c0	974	8ce3f350dfe12ca44f95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fd313416713bfea5fa80	974	53889c63251a67f47a2c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d7480332211f7ae0d50b	974	d131a3b33401a7b4a668	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0721e1d989e27939a324	974	d13bbcf819ba7aedcab8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ae411b1d56fbfbba6676	974	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5eb7191b60475a2adfa6	974	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ac92d01d621b00e3737a	974	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d466ac1f44fed9c3037f	974	8ecde88989703bbe399b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d56eaad238731792ed5a	974	a89ce690562f1bfb0bba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f8b77713b5c9e91d4322	1918	d8a47a56012378b52d48	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a288a139f32bd9566463	1918	cbbb379c9906538aaaea	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
596b9748d9aa65b80e7a	1918	4b963bec52944fc95745	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
578f39b13666a4382ce8	1918	af2892f8639e42fdc83d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
06b56b9278a270a8b198	1918	1b29574d86db3b63e11e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e810ff54306deb6c0535	1918	ff8a7abaa5467496e924	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
588f73b0bf5280841998	1918	7788942930b1d0ff75cf	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2bbcd7e1e1e2167ec656	1918	9ea43fbf389f497e0a95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1849a8b52c65031f0899	1918	a89ce690562f1bfb0bba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ef3a0d0474ecbdb8bc91	1918	f3f87a4f6848eef065be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5fe1975cfba3cf4e22ff	732	dfb67ad547da1c74dec0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a2840289624dcec28880	732	08cda36fbe812c092c9c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0f2cd856b01c7f02cfc7	732	f27cf82c1214ada6fb4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9d218711ad0d073ff969	732	ecf47a0bf5f29be17314	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
17c7a552824c165293a4	732	15341f322b282b495094	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4096ee2651360756edbe	732	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cde6a49c6fceaf961cd9	732	ff8a7abaa5467496e924	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
32826b0c740504d69f78	732	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
62c329c2b78ae5b7fffc	732	f212ea07533e401eb4df	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9d4c8923466f52864886	732	a89ce690562f1bfb0bba	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dae398b68a3744c400cb	302	cbbb379c9906538aaaea	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
83dc1c4e053d7cdfd632	302	02938d7b62625b948141	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8906e5219d05f9a56bc8	302	4b963bec52944fc95745	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
caeb1b1b3e54ce031722	302	af2892f8639e42fdc83d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
741c639756fe9747f0e2	302	1b29574d86db3b63e11e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e6da47c497517e6f6445	302	15341f322b282b495094	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0433e866feb209a42424	302	ff8a7abaa5467496e924	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fb3ffca75dfd030a9e0d	302	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
98d70cbf41034c4e92c9	302	9ea43fbf389f497e0a95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fddf59a8b5cd7c9c4192	302	6e57422643f6d74880c0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8c539c4aee55d50bb7d9	306	cbbb379c9906538aaaea	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bcbbef08703d368a2c14	306	6189d54b165dab8bd56e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4128f61b89564206557d	306	a82276e194b1732711b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c253127d70f2305fcaf8	306	f27cf82c1214ada6fb4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c2bdd560343097f2e653	306	ecf47a0bf5f29be17314	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
58fdfef828e64948dd0d	306	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e5932dc5622a9f859ee8	306	ffc06f66d5b1395c4fe6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5e65e714c01fc3e27d37	306	ede2ca1d19344a1b8138	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
42f07a914b10ef891760	306	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
60f49dea205cb707d3df	306	6e57422643f6d74880c0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
298730bb2223b6a91385	307	02938d7b62625b948141	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
571022709a7532b47870	307	6189d54b165dab8bd56e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9e73dd597ca424c6e714	307	a82276e194b1732711b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
92b936f0ccbf34a4c006	307	1b29574d86db3b63e11e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c0a19111da9f2db8c1a4	307	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
efb74f03830e90cf6dbe	307	ffc06f66d5b1395c4fe6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
95759ff38a2782bb0c16	307	9035d56968d798f556cb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
970359dfe43107be1d7c	307	7788942930b1d0ff75cf	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
57d61bb2d46ac6d4910a	307	591c7cad2de417c3536e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
14d1378b0086e421f1b4	307	6e57422643f6d74880c0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
24c049871c0f532dfee2	1914	98cc7263651f31776990	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5890b176add069109879	1914	366c6f405b1f8bc78439	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d8039609055e258d3b67	1914	08cda36fbe812c092c9c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9289041017286b3441bc	1914	d131a3b33401a7b4a668	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3878d5c852b688ab2ceb	1914	d13bbcf819ba7aedcab8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a0ffbede6aacfb39df25	1914	ff8a7abaa5467496e924	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b90b638ff493c5d84b23	1914	7788942930b1d0ff75cf	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
131fd16fd058cb41b6a1	1914	f212ea07533e401eb4df	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e06779de707665d76ac1	1914	6e57422643f6d74880c0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
513dd19fcb8b7be28ee9	1914	f3f87a4f6848eef065be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
07749e85e4cfaef5d083	316	98cc7263651f31776990	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
74f8303e19e0834a16e9	316	8ce3f350dfe12ca44f95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2290412bb4d9a62a30bf	316	f27cf82c1214ada6fb4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
624af0fc6c9878056a95	316	d131a3b33401a7b4a668	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9eb22098bfb1910eea85	316	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
997fe385b72ebaf72137	316	ffc06f66d5b1395c4fe6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b40d484d605a40d10884	316	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
45e7081ad46a429408c4	316	7788942930b1d0ff75cf	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
49db9e19061db2bbb297	316	8ecde88989703bbe399b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
63b61e030c4de0922a8e	316	6e57422643f6d74880c0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ecdf66c91e2309ddcd81	2394	366c6f405b1f8bc78439	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fddc02e5e9e6628d01e3	2394	4b963bec52944fc95745	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a32ec05dde1e77bc1d6d	2394	af2892f8639e42fdc83d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0c96217406567fd34409	2394	f27cf82c1214ada6fb4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2a0aa9bb4f48599ad19c	2394	ecf47a0bf5f29be17314	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
390ae858acab20f3d178	2394	ff8a7abaa5467496e924	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8a2706231035d9f7db7d	2394	7788942930b1d0ff75cf	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2c25c1b97207c14d9ebd	2394	591c7cad2de417c3536e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5af4851717af2de6c0fb	2394	6e57422643f6d74880c0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a705fba6b67f94a35aa8	2394	f3f87a4f6848eef065be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3bfdc0bb5941c68c2317	1840	98cc7263651f31776990	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2bbecce9afde0403a600	1840	8ce3f350dfe12ca44f95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8426e1f320b6a6784155	1840	53889c63251a67f47a2c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dc277d9500e8b4fa6633	1840	1b29574d86db3b63e11e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e348daab112879094f14	1840	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
50fd44be2807ca348ae3	1840	e295ccde749fd2982eeb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cdfa172aa1b868cf6362	1840	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
980ec955be430b0e9055	1840	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
67f9a682130b539e57e3	1840	8ecde88989703bbe399b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1a39feda93d509e79f2c	1840	6e57422643f6d74880c0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
952b1a1ecb1f21a4a208	1213	c6854f08ccb4b2c8da4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9f9bd9cc7f89fffefbf0	1213	98cc7263651f31776990	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2a2fbb027961fed46179	1213	8ce3f350dfe12ca44f95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
06a19e04061946ad0dc0	1213	f27cf82c1214ada6fb4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
987d47a385fce8b38490	1213	d131a3b33401a7b4a668	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c49d5fe71445d377d30e	1213	ffc06f66d5b1395c4fe6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dc4be4eb8ce8899441d8	1213	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c0d53d795ec1138fceec	1213	7788942930b1d0ff75cf	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
73835968da27b3ab2b69	1213	8ecde88989703bbe399b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
acc0956fc81c4c27ce0f	1213	6e57422643f6d74880c0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
992a61fdea4384784058	323	98cc7263651f31776990	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e6809487b0bb3c0f5fb1	323	366c6f405b1f8bc78439	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ffc92ce4ab65b12d0819	323	8ce3f350dfe12ca44f95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6a6a764ba67edcef82cc	323	53889c63251a67f47a2c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d3028497d62972e85a80	323	ecf47a0bf5f29be17314	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f1340217d3e28340a9bb	323	e295ccde749fd2982eeb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e4fa61af7e7a2442a4c3	323	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
75a849534ed7e5426180	323	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
adc952852acc45b6f220	323	8ecde88989703bbe399b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2418aebd8c9b93947436	323	6e57422643f6d74880c0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
176d4c40fa98e4f22870	2172	cbbb379c9906538aaaea	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
05ac41c53f30635fb97e	2172	6189d54b165dab8bd56e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
76edee3cac8f577de6ed	2172	a82276e194b1732711b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
952429f2aaed0792b27e	2172	0bfa86e7a44f47e58019	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
67f5c466f53543167589	2172	15341f322b282b495094	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8f002d470cdc8473aed2	2172	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d7ae62b314e5774d8160	2172	ede2ca1d19344a1b8138	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
89f80cfa63501a59bfc3	2172	7788942930b1d0ff75cf	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3a5d32f5afc0bad8e185	2172	6e57422643f6d74880c0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8b37affaa1cb8d567c4e	1730	98cc7263651f31776990	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fe5630ef92ca2b439e5d	1730	8ce3f350dfe12ca44f95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
935d4de7069703b08d60	1730	53889c63251a67f47a2c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
819f31e96c22fbc26aa1	1730	02938d7b62625b948141	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d06aef4f4c2e593bb2c9	1730	d131a3b33401a7b4a668	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
90fe9ace297d751e57ea	1730	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cc20bb311eaf1322fcf8	1730	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
95c82de7823deb9c97ce	1730	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
99360bfa5c1a3e74059c	1730	8ecde88989703bbe399b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
464cd85d3b7f738d5f0f	1730	6e57422643f6d74880c0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7723d4703369992c5176	958	c6854f08ccb4b2c8da4d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1c31d1161d39c9960d0a	958	98cc7263651f31776990	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5f314e1c9d27bcb228c7	958	08cda36fbe812c092c9c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a0ed4055e93e6bea9d7a	958	f27cf82c1214ada6fb4e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
80f62806c8d47b8858db	958	1b29574d86db3b63e11e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e6d7ad7ebc55f7026400	958	ffc06f66d5b1395c4fe6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
648db8b1254a954099d2	958	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5d7a9da5e3dfe2ca26a5	958	7788942930b1d0ff75cf	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8f013907df1b96da34e7	958	8ecde88989703bbe399b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
86af7a409decb8832e9c	958	6e57422643f6d74880c0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
50f75b2031800756885b	2165	cbbb379c9906538aaaea	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2ef0d7789cef4f8080cd	2165	02938d7b62625b948141	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f315f4bfe575f7106bdd	2165	6189d54b165dab8bd56e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7aec8a28c05eb56a6e2f	2165	a82276e194b1732711b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
99d4951549b51652a4ec	2165	0bfa86e7a44f47e58019	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
601d8ba7a5a3025796b9	2165	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
65acdc47d48c2490d63c	2165	ffc06f66d5b1395c4fe6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
de4cd797ab4c19a47cc2	2165	ff8a7abaa5467496e924	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aa5eea3e7c4b4a10ceca	2165	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d3df0d4190d36f124f69	2165	6e57422643f6d74880c0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
98cb1d3a810a4535101b	2182	98cc7263651f31776990	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
96d0dd9eeac14f4b1da6	2182	8ce3f350dfe12ca44f95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0d1ec550042681892515	2182	d131a3b33401a7b4a668	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b1332121564c7078f790	2182	4dbde706728519fe11b9	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6f86cffb4fd4d3e6f658	2182	e295ccde749fd2982eeb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c3a8154cf329c5850da0	2182	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ad06c906720a191ce46b	2182	7788942930b1d0ff75cf	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ebd271a55ea483e090d7	2182	8ecde88989703bbe399b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0fa3261638cd34f27323	2182	6e57422643f6d74880c0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
314a5aca844bb625e3d7	2182	f3f87a4f6848eef065be	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1cb52ef7eefb6584dd5c	1610	98cc7263651f31776990	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
565f8e1e61680971aceb	1610	8ce3f350dfe12ca44f95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fbc5ea6afa3a035fac1f	1610	1b29574d86db3b63e11e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9cfe1d19b00f2b275373	1610	e295ccde749fd2982eeb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b75e84933ac5d839527a	1610	96c9cb3e39287f99c251	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
71d68b78a699a2c1ad78	1610	1ef93a4c890e726a2aca	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8c9b529da3c45ceb63d7	1610	2a8c35f329bd530acc98	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
716b8547f9f967d9b647	1610	8ecde88989703bbe399b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fdf85ecd6ff5b609408a	1610	9ea43fbf389f497e0a95	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cc435f7fd3c543b69b42	1610	6e57422643f6d74880c0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d501b1ff3dd2867b6043	982	2399822855f02dd17bd4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
626372ae01bb78997280	982	7c1a4815aa1a2d69e547	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6a740f08fddb66902c5a	982	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e4c26380bdfb81bbbed5	982	a5012b7bbf099fe29fe8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a3353a21e5a3472ab53f	982	f6dcd719304f00d82c2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2a61580d63775201233d	982	b220f2ecaea0a0fb00dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
05c527107c4aabb8e30b	1930	7c1a4815aa1a2d69e547	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8eaa6d6bb3fab6487c4b	1930	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9dd0c2aced482e181f06	1930	a5012b7bbf099fe29fe8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3a87a46cf5e4dda02cea	1930	f6dcd719304f00d82c2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fe3ef34f5879e137651a	1930	b220f2ecaea0a0fb00dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bafd1bb8ff20938e7d26	363	7c1a4815aa1a2d69e547	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d9499ae14bb98f2ca575	363	ebfcbb48fbd63e18339e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e525cf3c967eb3b1db30	363	a5012b7bbf099fe29fe8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
491447eeee30b4a0ca0d	363	f6dcd719304f00d82c2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
84bfe1e97020925318ce	363	b220f2ecaea0a0fb00dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
045deb65f71007b52232	2580	f30291eeab3f00e3032f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2415dd0653d163931777	2580	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2698caa34b0d1618963f	2580	bb68063696175471a537	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6cac5c636d50bd611e49	2580	b220f2ecaea0a0fb00dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
01b1c496537f57e5c791	2580	1b092bfce595f85293ee	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c42c9dfa603f7215ff3d	1217	50a78e0f85245df78ebb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
81ea5846273d50cebb07	1217	0b72791b2dd715141534	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bb880642db5631faab13	1217	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e0a1f5dd220c1b041872	1217	bb68063696175471a537	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c83d4578345dddf82e9f	1217	b220f2ecaea0a0fb00dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d6f7fad75b8014818c0b	1466	50a78e0f85245df78ebb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
915889402b5d12b16bee	1466	46d83744c63768dc907e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ebf2c67190a569ef1454	1466	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
669d61b1c70fc20a6cca	1466	a5012b7bbf099fe29fe8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c46b18d9d9784f3c8df7	1466	f6dcd719304f00d82c2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
46732f06febbdaad6fd1	1466	b220f2ecaea0a0fb00dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f5568122963b44ac108a	2166	46d83744c63768dc907e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c4e9546321617d63e240	2166	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
201a4926524856226dbd	2166	a5012b7bbf099fe29fe8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5edd9dcc71b3deae1e49	2166	f6dcd719304f00d82c2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aa6298b0008ca015b9e2	2166	b220f2ecaea0a0fb00dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d1ead1cd35b1261fe31e	375	f30291eeab3f00e3032f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
375de4c8a7c0637d3ee5	375	ebfcbb48fbd63e18339e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2a84295dbaa960aa16e2	375	a5012b7bbf099fe29fe8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cb4c93a8909000001ec5	375	f6dcd719304f00d82c2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f3516933c0f4d937d3dc	375	b220f2ecaea0a0fb00dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9ed50b8264fe99f9e79b	380	46d83744c63768dc907e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ac90548911d3992dd992	380	ebfcbb48fbd63e18339e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d4ae749d9e35cb6ff7f2	380	f6dcd719304f00d82c2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7a2fdb3485ccde93ba9a	380	b220f2ecaea0a0fb00dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f1e1c4a8c0a1547bcf44	380	1b092bfce595f85293ee	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
115973649a689bcfddfb	384	f30291eeab3f00e3032f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4c5b72c61cd7f8b6744c	384	29e5c9ab72aa19cf5199	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9d41ce09efc1a16a700b	384	b220f2ecaea0a0fb00dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4f9831dceb49dad81e22	384	7e2d798388bf048f20f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
033513ecfa541a470f17	2589	2399822855f02dd17bd4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
546a2f73c4377fb8ba0f	2589	7c1a4815aa1a2d69e547	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
db36ce5f402eca0f1c7f	2589	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7b241e39f1b791b8f1ef	2589	b220f2ecaea0a0fb00dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7b555aa5760e79ae7ad3	2589	1b092bfce595f85293ee	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5b4b433a675662f5f40d	1761	d664df57dc8d387d1886	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5d042b0b8af7e461782e	1761	7f1e9e42b1894464421d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2f9588ef5c8f202a5bc7	1761	f6dcd719304f00d82c2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2a805c65f2f2c354b3ba	1761	b220f2ecaea0a0fb00dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e9500f05b0b53c03cdf0	1761	1b092bfce595f85293ee	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3d5afb30c6ab6dc2af34	387	2399822855f02dd17bd4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c98b0398099bf13081cf	387	f30291eeab3f00e3032f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d09baf606bf284a3c6b8	387	7c1a4815aa1a2d69e547	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
138c9cb60af5b5281adb	387	b220f2ecaea0a0fb00dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dafe14dbe0aa63024901	2167	2399822855f02dd17bd4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c80c09bf4477eedf1ca2	2167	7c1a4815aa1a2d69e547	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5273ff66da9b271e42bf	2167	ebfcbb48fbd63e18339e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dad5c450cd394ee5a8f0	2167	a5012b7bbf099fe29fe8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1783a6f5f0822aca42c1	2167	b220f2ecaea0a0fb00dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ffa20e4fbc544a8c43d4	393	0369837b7f5d361afb51	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
69b6082f38b61d7186bb	393	29e5c9ab72aa19cf5199	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c8b3b40537816d4685df	393	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
96ac506e89747737154d	393	b220f2ecaea0a0fb00dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
24b1cfda9dc370c9a7e0	393	1b092bfce595f85293ee	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bd8d3f50376593c873ce	1545	7c1a4815aa1a2d69e547	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3d0cbd751f9ffe8e6872	1545	242ae56440681672b491	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ab5fc932232d1614848e	1545	af736540ce33e44663db	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
24321da47d4cf9c1b276	1545	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
147ee755cb5e37d84939	1545	a5012b7bbf099fe29fe8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
006796ed73f8e63403e5	1545	f6dcd719304f00d82c2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
641dedd48b19165f8e0b	1545	b220f2ecaea0a0fb00dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a9b5ff9819cdbe7fce09	2603	2399822855f02dd17bd4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1be19d5da1db154bb3b4	2603	7c1a4815aa1a2d69e547	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c72049e860ac63241abe	2603	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8d58808f9e916ee23988	2603	f6dcd719304f00d82c2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2c19078ade8c2a486a4b	2603	b220f2ecaea0a0fb00dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
06d152da03620c6836f3	2603	1b092bfce595f85293ee	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e328316a2fa6378ab832	2590	29e5c9ab72aa19cf5199	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8393470ab82ec431f2a5	2590	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4aeb1ffa02e4a36ec488	2590	a5012b7bbf099fe29fe8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
28fef631530949cf7771	2590	f6dcd719304f00d82c2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3f93564edd95a43ed9ed	2590	a72e2dd120ef2c86ef04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9501bc2e95af99f305f9	2590	b220f2ecaea0a0fb00dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
db547302a96c33906dbc	1736	f30291eeab3f00e3032f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a327cc445711bcdb7d7e	1736	29e5c9ab72aa19cf5199	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2e07f66f4687c8b72d5a	1736	a72e2dd120ef2c86ef04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bb5ba7aceeb08d6e18d5	1736	b220f2ecaea0a0fb00dc	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bd7ba741e0500dbdefe1	1418	46d83744c63768dc907e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
72c2289b1a0f36fa2389	1418	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7504831e06b3237a9a2b	1418	a5012b7bbf099fe29fe8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f0407c58926771bd26b2	1418	f6dcd719304f00d82c2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
91669f0588a1b8e90d67	1418	6e2aeb5d8592447a31d6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
23dd473c2432c4c4b55f	1732	50a78e0f85245df78ebb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
057caa8903cc0d937b7f	1732	f30291eeab3f00e3032f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a8687bb5d104486545fc	1732	6e2aeb5d8592447a31d6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aea4f37ef12898cf0a03	1732	7e2d798388bf048f20f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
46ef4dcf8ae8c0e58b79	361	f30291eeab3f00e3032f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6ebed333ca589c5e2dec	361	ebfcbb48fbd63e18339e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b513d34d8cd123070da7	361	a72e2dd120ef2c86ef04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0b9bee260d31319162bd	361	6e2aeb5d8592447a31d6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
40111fa92526b20081ec	361	1b092bfce595f85293ee	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8881d25ef62780a633c9	1755	29e5c9ab72aa19cf5199	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b68c509639db8b61caff	1755	a72e2dd120ef2c86ef04	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6d9bc8d5c46395b47190	1755	bb68063696175471a537	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
079a4263bc666fe733e0	1755	6e2aeb5d8592447a31d6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9711d654ec058c8bf6e4	1755	7e2d798388bf048f20f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
af9fe7ff77e5c144c3e7	374	f30291eeab3f00e3032f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0024bce20b8a028f54a7	374	242ae56440681672b491	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5e58c7aa597dacb1268c	374	af736540ce33e44663db	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8c3eacbeee8e5ba63b22	374	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
53fc70b4690fd06f4eff	374	a5012b7bbf099fe29fe8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
be685772ba1bd1ba3a7f	374	f6dcd719304f00d82c2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c7cf0276703efe94d3fc	374	6e2aeb5d8592447a31d6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f822f119d5bec3ae9d0c	1227	d664df57dc8d387d1886	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
52236b5db2bb0c88cbf5	1227	46d83744c63768dc907e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
801166485d6149812c3f	1227	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7c8c4c42aa9be51de455	1227	6e2aeb5d8592447a31d6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3e3f24fe1e4ce464b59c	1227	1b092bfce595f85293ee	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bfa54292510b3be840ef	1222	7c1a4815aa1a2d69e547	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
61459c587d7a58d8c775	1222	242ae56440681672b491	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
601878944fbae29c07bc	1222	af736540ce33e44663db	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f557074df24de6ab4cfc	1222	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0f53a5861b9e571e62c4	1222	a5012b7bbf099fe29fe8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4ca904eafe9ae75aa21a	1222	f6dcd719304f00d82c2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dca916afb3fe8f7a4fbb	1222	6e2aeb5d8592447a31d6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8d1c77a6931f69641f1b	383	7c1a4815aa1a2d69e547	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
852871afb78f162fffd6	383	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
36b84869820fc4f8725e	383	a5012b7bbf099fe29fe8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d8eb65c9fa0dad37940e	383	f6dcd719304f00d82c2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b4b43e082d04368ab32c	383	6e2aeb5d8592447a31d6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3467ddd052780ba58870	909	2399822855f02dd17bd4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bfdbab624a47b077c92b	909	7c1a4815aa1a2d69e547	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bdc76285ec58b74a0929	909	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
905fe1e71dbb50ce61e2	909	a5012b7bbf099fe29fe8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a6cd056fc2396c8cbc96	909	f6dcd719304f00d82c2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e35db0fcc48e49315c16	909	6e2aeb5d8592447a31d6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
32bb8d52b3231efdd6f1	386	50a78e0f85245df78ebb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4cc1672dc6f8d6945f9d	386	46d83744c63768dc907e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
963c8f7136cb967d37eb	386	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cf802b1b33665630c262	386	6e2aeb5d8592447a31d6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4d646072a9559915d513	386	1b092bfce595f85293ee	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
11825efe11e93bac8228	2168	f30291eeab3f00e3032f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9f30387638b79a7a1b6c	2168	ebfcbb48fbd63e18339e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9062aa69fb49be11db8e	2168	a5012b7bbf099fe29fe8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
962c1b089df578386e0b	2168	f6dcd719304f00d82c2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b7d1fbd04f0f7a974b87	2168	6e2aeb5d8592447a31d6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9e4208e0f549e9aa43ee	395	0369837b7f5d361afb51	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1ec6bbc23bb12ca30ed3	395	0b72791b2dd715141534	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d8330ed9dc2129a68f4a	395	29e5c9ab72aa19cf5199	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
450e17fa079b53be44f4	395	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
44e2d3ab38073a76568b	395	6e2aeb5d8592447a31d6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
56f3f28d48039fbb660f	395	7e2d798388bf048f20f6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0667bf7d384ddbb9d4a3	396	50a78e0f85245df78ebb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c0839525dd7531497d41	396	46d83744c63768dc907e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2d9322cfdf6896b0ba65	396	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
33d4b6b1f5a86850614d	396	6e2aeb5d8592447a31d6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
edd9a7acf719d75cf87a	396	1b092bfce595f85293ee	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
707654f7e8de1d53605c	1467	0369837b7f5d361afb51	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e17f93ad3222926a4422	1467	2399822855f02dd17bd4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7987771ae0e4d00af4b6	1467	7c1a4815aa1a2d69e547	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2949f0177cd65d56b89d	1467	ebfcbb48fbd63e18339e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
0b1fe372473f9787b36c	1467	6e2aeb5d8592447a31d6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5b2e82d4bece40baa39f	1467	1b092bfce595f85293ee	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e20f8227be7798318f90	398	46d83744c63768dc907e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e16302bc0eecba906697	398	7f1e9e42b1894464421d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c5ea08022dc4b8115aba	398	f6dcd719304f00d82c2e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
795d30b36ee846168d35	398	6e2aeb5d8592447a31d6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4553d1e4e3fe482fc731	398	1b092bfce595f85293ee	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8f786f309c0dc257b192	1735	0369837b7f5d361afb51	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
82e450db495762508dd1	1735	46d83744c63768dc907e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d6fb94dfa6b470739de0	1735	29e5c9ab72aa19cf5199	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bae5203ce80a8a937b2f	1735	bb68063696175471a537	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e964b90ad60f02ddb19b	1735	6e2aeb5d8592447a31d6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f182ef937a7c3b817be8	402	d664df57dc8d387d1886	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c4a35d25c0c78d941460	402	46d83744c63768dc907e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
93432132275401c2ae02	402	ebfcbb48fbd63e18339e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
70f71b3e4fd8924b14b9	402	a5012b7bbf099fe29fe8	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9e86b7b3998f189eb2ab	402	6e2aeb5d8592447a31d6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dc46d9198cee2845c579	413	50a78e0f85245df78ebb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
19ad28e39b89232e4d65	413	46d83744c63768dc907e	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9a2407731a9ca9f25c16	413	7f1e9e42b1894464421d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
72bfe122f9a5d0307b94	413	6e2aeb5d8592447a31d6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
447ebb9a507957c0f947	413	1b092bfce595f85293ee	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a546103363a7a9be7974	1737	2399822855f02dd17bd4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f6dec0173516bd911996	1737	29e5c9ab72aa19cf5199	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
63d3ba80818d229aec01	1737	bb68063696175471a537	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3f431748d746ba39f6bd	1737	6e2aeb5d8592447a31d6	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9f116db17e322d70b2d7	1974	50a78e0f85245df78ebb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
fe103bd93a028f4485ba	1974	c7844f6432c9b3505784	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
db586ae4550ae16fd453	1974	e6e6b9eaef349d8f67bb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
67c303be559c5dd206dc	1974	390155dc1ec0a20487bd	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3b5f42d0f8274d3583ff	1974	78a0f435af1bd5acb2c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
94ccc67bb037b5ccc4d0	880	c3c9a4fcec3d5a4cc8f5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7c1b84972cc4193258b4	880	e6e6b9eaef349d8f67bb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
55ed844998916bd47182	880	5d0f2a46fc3857533f25	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c3c9da9d99088bc4a777	880	390155dc1ec0a20487bd	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3f431b941d0becf19190	880	78a0f435af1bd5acb2c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ebe60bbb266a301fd11c	735	eb6a091792f47c508880	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee424a97de51fe12e0e7	735	e6e6b9eaef349d8f67bb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3539c90ba8a6bebd7531	735	2aba875ca65d820f85b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6dc411b8785835f8a50f	735	30720b1ba7b7ea205543	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
623d3c5ec8a7a530f73b	735	78a0f435af1bd5acb2c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
765166a108ffa90f77b1	427	d664df57dc8d387d1886	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
354778193d8f7b32c999	427	e6e6b9eaef349d8f67bb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1757fd9f06040f02fa11	427	5d0f2a46fc3857533f25	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bdb3f66c65ddd5fe358c	427	390155dc1ec0a20487bd	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1582970a1cb48d3ad346	427	78a0f435af1bd5acb2c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ba7b884a375bda66ead1	737	0369837b7f5d361afb51	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
23e1788d9c6ed7dd1d3b	737	d664df57dc8d387d1886	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d61e19a739f0d6c2f281	737	5d0f2a46fc3857533f25	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9e371039dcc623458be3	737	78a0f435af1bd5acb2c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9d215725a6e327eb2b0c	436	eb6a091792f47c508880	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e9e41395677623502dbf	436	c7844f6432c9b3505784	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
811bed325c08030ba093	436	00a45de906cbd0942534	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f118d175d5101cce4cb1	436	78a0f435af1bd5acb2c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
76156dafcc16a5063a7c	981	d664df57dc8d387d1886	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2fcc4d363970fb7d6536	981	f30291eeab3f00e3032f	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bb647e234460d0f19370	981	86f346656a8fafa3495b	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
74058c465eb814ce3547	981	78a0f435af1bd5acb2c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dcf87feb9e49f7329f47	981	1b092bfce595f85293ee	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
84f2038acb04d2f308e0	460	c3c9a4fcec3d5a4cc8f5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
6a2d7d3ad9ed984dba03	460	2aba875ca65d820f85b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
16098f215e97ec814915	460	30720b1ba7b7ea205543	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d34eb8d53c43a9a4c60e	460	78a0f435af1bd5acb2c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
35cde4251cf42ef719e3	460	552ab552eb8762c570a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
20e090a373bbbfe79421	523	c7844f6432c9b3505784	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5d70346119e795eefd44	523	e6e6b9eaef349d8f67bb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
902a0186475d1d6c1d41	523	2aba875ca65d820f85b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c924e4e3f1d9c02f9a32	523	30720b1ba7b7ea205543	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
85770dc8b9ec4bd94f58	523	78a0f435af1bd5acb2c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
015c245e469e53172e16	468	c7844f6432c9b3505784	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c4f2740ce609f7cd2aec	468	75a0ab6961da49ab3b89	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a6f3547c41cb8e20957f	468	2aba875ca65d820f85b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2cf4e0448e9843db4ed6	468	78a0f435af1bd5acb2c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
41d6d451c84dc275d516	468	552ab552eb8762c570a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
50089da49220226e61d2	1472	c7844f6432c9b3505784	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
acc9a1cc57bacbcf2f5a	1472	e6e6b9eaef349d8f67bb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d9e7fa343d165f7a26f9	1472	2aba875ca65d820f85b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ec468fd8f93f5fa74934	1472	30720b1ba7b7ea205543	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
165867d93e67f650a524	1472	78a0f435af1bd5acb2c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
60d7510d38e52cf61df4	1794	5d0f2a46fc3857533f25	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee33119c2f83bad91f15	1794	b29c20c00222a6c1876c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f51f56e803057d59d1cb	1794	f75e70904a70afafbbd1	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8a09bd0790334dab3e27	1794	78a0f435af1bd5acb2c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d8a1f4f592381ba8d713	473	eb6a091792f47c508880	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7bbd1f8a6f3202be3f16	473	c7844f6432c9b3505784	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
57fff1b11111913e3dc2	473	e6e6b9eaef349d8f67bb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
038f6b3d62f441631955	473	2aba875ca65d820f85b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
136e6f288a5394e4800f	473	78a0f435af1bd5acb2c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f2cd224b01d40a4f5672	474	c7844f6432c9b3505784	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
aa4cb5b83279a8083d1c	474	e6e6b9eaef349d8f67bb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
450da2542d65a9db2007	474	2aba875ca65d820f85b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1091838f41aa5060a7c5	474	30720b1ba7b7ea205543	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
54407f4a469400d6cb94	474	78a0f435af1bd5acb2c7	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
5c0c71e3d5f69fcb0499	420	eb6a091792f47c508880	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a14ebcc5565ccd9f9dff	420	c7844f6432c9b3505784	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8f579ab811ecebff7a95	420	2aba875ca65d820f85b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
95798f156499ea3a8669	420	9bbe8ed0f5c6182ff324	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c4ee768796e14f963113	420	552ab552eb8762c570a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c4d261289b7a90d8d684	1738	eb6a091792f47c508880	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f8fdb58e7de9ca8792fb	1738	c7844f6432c9b3505784	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
250674444926c60bd787	1738	e6e6b9eaef349d8f67bb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
a70100a574474709ec43	1738	390155dc1ec0a20487bd	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
2b42913b18c42daf47c0	1738	9bbe8ed0f5c6182ff324	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
71be9e2210e40851844c	2292	eb6a091792f47c508880	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bd660f8d2b0f953aa6ee	2292	c7844f6432c9b3505784	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
36785a5bd56041265c0d	2292	2aba875ca65d820f85b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
bb45547ac36eb82497ce	2292	9bbe8ed0f5c6182ff324	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ed489ad133b13c51355c	2292	552ab552eb8762c570a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
7a4dc06246c2f4a61e62	1926	eb6a091792f47c508880	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b84a7b08a14e7a8f69b7	1926	5d0f2a46fc3857533f25	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
cc2852a7ad8944246d37	1926	00a45de906cbd0942534	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e266939b70ebd6eb49ef	1926	9bbe8ed0f5c6182ff324	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
78ce8788ca832724ead7	1598	e6e6b9eaef349d8f67bb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ae4fa37eecddb1eb6245	1598	390155dc1ec0a20487bd	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f94ebbc0e64f8f74c8b2	1598	30720b1ba7b7ea205543	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
9a52326fc629fef55f17	1598	00a45de906cbd0942534	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3da1a7549068b4806f46	1598	9bbe8ed0f5c6182ff324	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
e19cf354be08ece97973	448	0369837b7f5d361afb51	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
385850c94db4ae6f75cc	448	5d0f2a46fc3857533f25	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1dfb18161d2df9c82a9f	448	00a45de906cbd0942534	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c783dac3210e832b2202	448	9bbe8ed0f5c6182ff324	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee52c020a4b3a3c0f2df	1471	eb6a091792f47c508880	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
191a206b6565c872203c	1471	c7844f6432c9b3505784	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c0d3e313be093583363c	1471	b29c20c00222a6c1876c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
b00b496484cf29c6d98e	1471	9bbe8ed0f5c6182ff324	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
793978593ccea48591de	1928	c3c9a4fcec3d5a4cc8f5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
8dcdcb1e756aa5d7748b	1928	e6e6b9eaef349d8f67bb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
51147f4928cd2356db4e	1928	b29c20c00222a6c1876c	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ea206e9498a6abac66e8	1928	2aba875ca65d820f85b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
d159e89e43aa568d14a5	1928	9bbe8ed0f5c6182ff324	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
51fea4a983611ba694c7	1561	eb6a091792f47c508880	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
07c63a3cb51d6d6e3e6b	1561	c7844f6432c9b3505784	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
33cff24add80840f9c7d	1561	2aba875ca65d820f85b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
93764210486208dab20f	1561	30720b1ba7b7ea205543	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
c50d6b7845f70353c476	1561	9bbe8ed0f5c6182ff324	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
abee0d4471c5358dce5b	1561	552ab552eb8762c570a0	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
dcbc6f2f77822dc63c28	1184	0369837b7f5d361afb51	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
746df8b3347e2e7830ee	1184	e6e6b9eaef349d8f67bb	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ef1ec1cdddd48d12e909	1184	2aba875ca65d820f85b4	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
3cc4e7e3a9cf2a19f288	1184	30720b1ba7b7ea205543	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
79fb8066a87ad2b52b0b	1184	9bbe8ed0f5c6182ff324	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
f63b88a496d94c5635c8	976	c3c9a4fcec3d5a4cc8f5	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
1e5d91e0abd2fab1e556	976	5d0f2a46fc3857533f25	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
92c0e238e88195715af1	976	7f1e9e42b1894464421d	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
4fa7c6ac29b783abaa5d	976	9bbe8ed0f5c6182ff324	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
ee6740cdded80163d5be	976	1b092bfce595f85293ee	\N	\N	\N	\N	\N	not_required	\N	\N	\N	\N
\.


--
-- Data for Name: AuditLog; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."AuditLog" ("id", "userId", "username", "action", "entityType", "entityId", "details", "ipAddress", "userAgent", "createdAt") FROM stdin;
pm3p55jscvk32cinzgxqmk7u	user_admin_001	admin	create_user	user	v152mbwu3vdblom8o3xwoad5	5505aa3c9f3fa59c12c0d42f:ec31acadff988c799919477d0c1d1c7b:43b29be70e8fb6a137ab97c2e76083ae610e8360d3d1dd8da8a9d06fdfe6b291b3dbd8f2f3c350df2072c736f78b6085	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 13:20:54.901
lu2uk27tk3csn8m6ggvnw55i	user_admin_001	admin	update_user_roles	user	v152mbwu3vdblom8o3xwoad5	4badb1064791bcececf541a9:c244c9ded5f04e0e70098377e8b21708:fcc2b7dfc83f605591ba2003db9abf1fe425d7d310da204269f843e97e118c3bf0612e0263bb67b24d92734cce351c521583e30a56d8b842bfbdce37d6baa6e2f3d3be4297d1b3194146966cdd6a8dbb605829c1274d68ea	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 13:21:00.482
mk41mlgbco9g7zcz1ku304bj	user_admin_001	admin	create_subject	subject	DT	aabd9f37fb2a0abf395c6000:12e47aecf948e29e87272b686a9f9f4d:fe76800558cf0c6cb9fb7cd374524614672ec5a77a2e1eee5b9a58ad1705b4fc6382713ec64b71c8592f80e8d571cbde87400a20ec75dfbc7ec7ee	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 13:43:02.248
fbf9xfcpoceg5gr0n3eydqk1	user_admin_001	admin	update_subject	subject	DT	242db115e2c26dd5d82ae88f:a33583bebce7e4d90cba068f60dd9e1a:4887d9362275af2bf565bf8a627f805b2fcfaae895e6d46a7420c77d10b38a756bc528f48a8a70a9157dac21bdc437dfc7799b71effe7dd431dcbd8b99c17944ce294065ba5756436438e9fdc7770bc8bf2d764d1ced91e2313958189a8beae283462e5139a271a9036fa5b94582e6f5f31c82513e7bfc928149	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 13:43:20.712
v87xkaxp3vg9m2a35q1ealzn	user_admin_001	admin	create_class	class	d9j84yuyax8dyz050pzas0zx	802687e9d233121d72652d1b:bf27c0aae6fd81296f08e1ccc5aad10a:1c08b9597752e7423a13a1c2f64b235ea0207b3f7d384e255c280a7648ee97e2c6db18785f2bc58a0c6501614b0802a47e80fa4e9e34ce5c9bb0a1730aabdc0d8c8115bcfb0f9381c3b99281555ef060ba9ee9b54358c9716b6dd09a21f42fb299	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 13:43:34.947
nrlmfqkuu550w2g8ieil4q57	user_admin_001	admin	update_class	class	d9j84yuyax8dyz050pzas0zx	05398f7db21ed952fbdc0b58:9639184f1410b6d0cb23bf569ebf1270:2fb0604862efbee0f9b6ad7018034d221b2d9caf2c5dc0733f687a92d8e34fa41e64d646ee0b9ed0be951ef77d96b0ed6023b475266cf24637b44817dca410910b404f306f65cd1836425a25f23b3027869bba3246aa54c2400cae97f56cce63b742c249170371072efaac5e1d81	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 13:43:40.525
n12b220qcbyhj78ktkstwv4v	user_admin_001	admin	create_common_comment_group	common_comment_group	kvyc23qpstalii62z2ytnrl5	00c91ec7bf83633a61029568:251d7cc9cdc1abceaf4e4897ff14e5fe:d5d02cb13d989114b51da5eea212fff02f7a850898172763707dfee45fd30a79fbe7339c5019cdf6a25787058dba7dc468f09ea7d44eda8201641cce4ea817a6e015441e43cd827da391c6f594afe60d5681ce9fd67beb0d4824463f98bedefa62	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 13:44:30.439
uw5thudpt5va9eqkkf7kwl69	user_admin_001	admin	update_common_comment_group	common_comment_group	kvyc23qpstalii62z2ytnrl5	0b2ad6102eef1095952a2bb0:670751251ca042834c4d4337e78809b8:e399fdeef51f2a6dc747b19fa29bbd09725a5d9a00493e4dfc57ce21710fcf2548b259df2a70154c18b1107f09483429be2e5ff3c450b69302c053a2971fce3da95cbb9786e765f01c9a7f80dd761b1d5c313ae6cbdca48fcf0f51b2d8c5ba387af9bbf675b2caced5e5cf8a5c2717180332304a858a376759ab2e0480023140e160b590dcfc72dd9c5a148e225c7e2d1ca5b55ca424ad82d3b4181e3231b87d6f6eba2c758bf05e8bf85648dc52995f5faea7cc3049d8787ab79734c2afe1	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 18:01:59.964
gj4ax49k1o8h7akz18ho7840	user_admin_001	admin	create_common_comment_option	common_comment_option	r3rga6rgas2akmt7m6i4ktok	e393e682962f8a12f6996a19:5a93d94e673c542c39f9166a21c16267:365876afa5834cd0aa5f17902bbb5b8f995847bd0c911a9741094ce43ad685efdb5bf3eceb17acfa720539ee37c49b3b8a2fdb03c7ca6dfaefe2a9da1dc2b0b587d3963b58309ca7981a22a5	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 18:02:13.178
gz5tvzuz0cne30a5ofl9hhpr	user_admin_001	admin	create_common_comment_option	common_comment_option	ga18zj0kz662x618tabt6813	6a5b0a209016a2851fd3e4d7:02caa108a83031edd3c53997c051e9d1:914d77db371aa60fef88bfc5be84c8320ead40a9868c662b97f22d8245d5c7ae72f97b7a9cffa7b1378c151b461bbb1f8c610a863efbd69e94e5a396b7b686cc495683876af18d40	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 18:02:25.344
x26f3n3uwnuvim0l53f0d532	user_admin_001	admin	update_common_comment_option	common_comment_option	ga18zj0kz662x618tabt6813	6a9b320fe02c034e90bf02bf:07ab6b2e15112c8af1d041d7d636403a:108fc03b29d46e536a8da3401689ae846c5fb996630fda75f713df86c3ffa165ae948e87eaefc8c5f99cd0ea4cf588bb66524af336fb2da08bd0a68cc5a4767509d54f7a8edced0c2b8d	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 18:02:33.725
z8y4hjt023a5objdpo4yj61h	user_admin_001	admin	create_common_comment_option	common_comment_option	t3n6ux6qdio6uxkigk8y67x5	0234625f8e899b670a21d13a:42ccf911ea5c694573fd25aa737a0d07:6bd16ecfc476ef1f9d0338068caa94c9cb2fcb2116de6639914c20c3745e4ea0c50fbdeecb3b794e4c23e316e3f43873e7abd53ef000b32ec5fcaface327f5ebe91e0ff804ecc9c1ece3a8a4	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 18:02:49.187
ym03jkrxpok53cneeci9ita4	user_admin_001	admin	create_common_comment_group	common_comment_group	m7k11htb7ybhjbzra6a0pbx8	89000932628b4269ff72d1ca:4aa07efcd59b79b6a94bf7cab9fdf391:783b65b929529eb9b698aa0ab48d66778af16a60b1c7a780db7d07dec79da3ca9e7193e012fae084d1928a5d27fda9d5533b14e09b1f00c9297d6c249bc565c8352cbffbe88b60bbcd58ac314bad4a89a62094d8eb665b5ef3dd7a770e8a1009	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 18:03:36.475
qfc3bsw8zx5ni4e1s5sgh7c5	user_admin_001	admin	create_common_comment_option	common_comment_option	cs3r0ad7jwxtv9l4dfvvukce	528cacecd21557ea5f7b1e08:703815d2790f222b2dc03d296b768a64:4af0df35f66b359b5a5ca50f0ea37dea87b24d1a2eb46bde6fbf02424e45066478fffa0efa465e1233701e598cbde7b18ac06978ff528475a2b51ac4446783fec89efae686d453877a2b38db159993b1fcbe3ed645362d422f311802f4a72766832c82a9c399c17413e51695a3786fdb5587eb87a01b3b4c12164eac3dc543326515bf53ea343cca6d4317662d02f31db045be46007d8be915ca75a64ab6af6a04a17bca59869ee623a6c3cc47e20f	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 18:04:48.929
eep7jhom550kvbvcevmd2d8n	user_admin_001	admin	create_common_comment_group	common_comment_group	gq2e46phsivlcljdybugoaw6	0220fc1d0e50072645ecab28:6dc6a659d0f8ade85c54aa96dc371980:891a1b4e813efbb1369d6e62f032b399a9b335aed1ff41825922bbe41af42c9950e27f45db15e78f4fad63abc72103adc236cf77007ec2761b84903c78aa5acedd23b8a07cc6ebe89574eebb71b2cb10d7578c5c1dc0	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 18:05:38.778
xxb0kihlsrumpbiobb3o24k9	user_admin_001	admin	create_common_comment_group	common_comment_group	fht0kqpvp21kxsfwq6hyalbm	72022411f35bd846731cf139:fb22053c9f1a181801cb6c7fc3bf4b80:e3b56f72bcb7c92ff8110eae143ed990f5a5a90490f638f985ac1f34771ed352212ad189ac18cd4450f124652c1c094300b3542d83c76c9f2cede54c125a0d6d98a36510f68b82cc305aa823e013b08ae0e00cd8fae5619290a29a057f	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 18:05:54.369
yosv2rv3xq4numameh4q8kvb	user_admin_001	admin	create_common_comment_group	common_comment_group	n5sy0axnb4oj9k0nvnyxkc3o	44d1b850f9f729d4020085bf:9a0b9ab1eb7956ac58d068e2840eff7a:4656cfc144f06fc87d8fcfaafb37e718c17ad149d5a2f7d7f3a53aa4918cfd2d89fd7746ebe7dd75bdb2441febda4216b04c62f3df305a617384389551f6ab61f9bc2c14a2417f203166984f4b899f2a1dd8f9cb27d641e2b1ca211adfbb	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 18:06:17.004
q33qn22856l7ptmtaekcm9jw	user_admin_001	admin	update_common_comment_group	common_comment_group	n5sy0axnb4oj9k0nvnyxkc3o	18c8feb3ca415deb981debb3:1a7b9e319991f98ea461fe1955a497b6:f2b46978777af2dd2a1c11e41475999fd021ae7531ed17e6e1d6d729cf7f4aff6808007f6d47e3c7e5a544646d046175d92333a852dcb13713fa14d91910239edc4f2532c8fef61e50083de52b161b51f528c0ac8ae6a53b334b3dec7f744a5ee44fdbb7dd4d5fa0826260ffafe71938367e3659ce581f5562419339d1a73dabb516dc99a1381fec0add6c23c742ed9000df6f13f468692675ce7e77148889c27e042b159aa87a911b2d81216d8a9726777ac9c7c8d6669fc550edd0	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 18:06:29.234
j20g1hsj26q5urkfxs0ev3cr	user_admin_001	admin	create_common_comment_group	common_comment_group	als7ui0tyuuos8j58jndt892	da0e49e7f89d7e88c53a32cd:87551d1ecb41c279cf54bbc34195c967:e5e9f4ecc82526e519408f5905a433144875e2892599f3298ac456a8cf4a974e4d87c1e5b6f38f9ac291c332f5be6455c83eabe24f79b23801255e841d4792af3be823252cb546c00107aec9b080ac664edad46eb9f8c46a3e188a920930f633785f50	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 18:07:25.969
hkqmhf2shrpwkfzaoe0cd1yh	user_admin_001	admin	create_common_comment_group	common_comment_group	ismry4x2kzdwssn2pdbkbu8f	3a9938d5097f59aaa02a4fb5:16d7aaed54651396db49419c540d73ad:68a3ce81c8b8768312f320777a156798a07754e6297c168ee8f86b7bfd93b2aa0302b3ee7cd2d2e55af13fcd0fd14e9ddc2a67c0945eb56169095c4a5adeb53c33bb05ed0e35c41f839a0ab6aaeefe435e24	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-01 18:09:22.752
s3afywafwkj3xwjc6gb8ubed	user_admin_001	admin	create_comment_group	comment_group	sdt5wca5bot6d8lo2u3n61yx	ecdbd4cd59493ca03849eaea:9102c09e293383915dc0047fe5fc61e7:376cd1e377cd4e36447e60b6c4d62c96ca344bbc98fa7e936b0e227caf519ec80b5764fe589e14f04a7380a6602e97a4eb254f69a9751d8ad34a8c8001f7f11499cefd0951992ee4890c9f5ce16724c6bc04c4ce90a838b9993a3f29718ff550c7ce54fbda7e00b8356a8a6f07f47efa1858e90ab11e72b8a2e2faba1893dc	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 09:23:42.146
n9ze8jk7w2p428vc8t1i7ty8	user_admin_001	admin	create_comment_group	comment_group	md1b4mj0f6ztl29ksjkani72	6a0cfc2e24f97d79a7cdad49:ddf12e3cd0beea1541702e2973fa4cb5:d7b9c08f7db9957cd92bab86755f631d832cc93d7c178cc2c8065e7712f924e02f12cb3abcd1633eda293e4791f9211d4c4db1e05816cc98aa7d19f8ae31d124b5c12a88d1203b6915284c98ada1b7389894cbfc78e592f7ad8e032a4e4c4cde14c821b6f4b48a73765ea041b33d	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 09:23:59.005
fhbgl2npicbflnnr6djldmdg	user_admin_001	admin	create_comment_group	comment_group	yb8henvd4z79t4xljb1ciu3q	9db4172509aa3e0be02882a8:0eeca1d23e22947ebc2c93af12078c7f:7cd7f97176e9d7e6bbea6b2a9b164ac36e6802beb51952f9abd4569e31c3ce110b90d839cc341dd2eab7e5f69ff2968347b8f69e4a38994e7ac5c7afa1d2b86cce69e790e52cf2a1ddbe265d8249f77a934b76fd9f288976f7ec697ee6b9c9c63e0720614964684baa72a0ca2b24df95d4	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 09:24:22.124
r1bhzk6evvazdiw6r3x6f50z	user_admin_001	admin	create_comment_option	comment_option	f1x5gjx4qu7q1d9n3t93pseb	585349ae969ab6b47cd80dd8:48bb4ae4b0a75a88baef5fd10dca9d9c:ac00ab218207591a50c1fc36d4e064390bf6fd586301d388a7a4ff81a6f5f78d9261e70444b45643fe8cb81f20308db2db52749463bd9d4961a11359510299ff5c210ed4bd563bb20409086dc153136b054bc9253968d691fe71f6a7a7112a8105ec0b523b4674bcf7a1546f2fc737f695428063839a55abd9fe4f86681f420fe753021146c25796b92b8860c327b9dc84a4ea234c118fba4b8314301789b9ae7c6a01c224d4a0466c76992988db21824e068e5cb83b9c07f7a785f99840c3d8f4b555dc78ab956f291b478d94d4318ee65f1d981a3e8a67918ada47d12ddda34a7477ece4fffde5ce5da4642473de855b929ed5b438f8f6719dbdc160d4d120045b7000be1b2de697edfdacb4f4ebb4e94f50f5966e9e1a6abfb171293951e881f067e4622a9c3a71eb2fcd5103d2ee88625f2a5d7c0c3bb946e9be9babe54864018481865a857c624a6dfa2ec30eed48fb8bc3abb720b925ada539749dbb9febae7c657b483d1edcb55376502f95b73d2163b27fa5309b937bf597bdbc05c50173dcbfd0357551537f6e4fb6b7f6890542d3d98e23c8dda6d37ee8c9709cd3c80306588ec5c8d522ec26c373b3ee2773260146f21b660bb719c1345f03eea2977e9ceef28be6e3b47b5e919a63cf81f70de4ec8eaaca3090d6befd4705131a6e84	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:42:22.269
krp91g1d7vn2edri1kjombi0	user_admin_001	admin	create_comment_option	comment_option	qaz0yalxourn46tzosgcax4s	e775cc1a6ab7b7e01701b814:773ad25205fdfc6cbf76335b8aff1c00:8b292aa1209c3a8c125784aac5d9befb2f9347bb469b20f5990d149712bb0626aafb661434c74626a4063645c53214c3ca943033f89588dd8a54b1c931beb593193a0d44e9d24624e5254c68dec5fa584c5cecf9685263af871315d703876857131799d506e8fc7a786b47813195d50272d2fd376816abcfddd9ece7c041c39344ec30c123dc349ad36438ee85eb90e6da215ae4055e454a946384682b882e8857bb22f5768872ef59548c220d653ad65fa28d4dcb44a1ff9a8116e6bee7895d0724e14d2456f1aba022aa73aa79b124fb12e5069cde998b05363ebec2cbca007b54e075ac16a55eb17556d0f32154a57f9766fc911f251bc0d5aa4e180bb7bbc5f6ccafbf7438fd0bb31ecc3e411a49728bf79494ebf3fbc46f6decaea9cd62c89d538249b2e7d9e99fdb059fcd56126c4b8dc5e9a48d0b3550fa646d97975b52ae14bcb1487a681fd54608cbda9662be23c42fc5cb6ea9ea92f76afdf08ca10055e4fdb70cbdd10bdcce83e9f12d4cd4aa74acae1474803ca5f896f761c4aca13db097fb598f5ee30b5f92703fbc271e436e34e58b1f2c00270d26836c3606ed4518f28568cfd618d71b8e0a592b	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:42:31.923
iz3q2fwh1fvsmlpqmqhjzdrd	user_admin_001	admin	update_common_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	6f3e9c1b60d7bab09f7e01c5:a845c4ec23f63b254bdbc0e3e47f240b:23364eb9c0c38bf565f35e2bd9b516d6a3da80879cb5354570c1ebea5c8f9d0868c8a897cf16b28d05a75c2e9843edf9156c690df39691f7eccea5ca85d06c3f00dd87e99aafe7a8f69083300b0204f622ef958645939c1d6f16d8ef5bfe156654f5ad3c92acbe6a557433f220220779a4218f66b601ca24c615e89d69212238c24442e60541	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 07:56:07.939
lqvx1il93geu7ryxqe6yd3p8	user_admin_001	admin	sign_in	user	user_admin_001	5b759ea1976a96d0f7e6a209:e952804c445b9d5a591863a61ec5bce2:455b49aca994b67d408f69adecdee5759d19a85461dc5568e26002d88eca27ec020381	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 13:40:18.887
nmrrebosmw8erx9sq01f5e9e	user_admin_001	admin	create_comment_option	comment_option	t5hsops220j5bu4sm3ikmoc3	bbc6e63c07787cbe73052a83:c028ad8fdc090963d4a69b805f1c040a:43b8019fcaa9097f6859ba1805d5f85de0796799ad0ee271b06e714424c51cd2efaceb2fc3c1273363bd974918076386bc43eedc8e652987cc52cc84386643aa37bd0718abf02ecb2e298b60898705282303b4ffd444cabab0be44144ae0bc9f196f24cdb10db7432780bfc6255e316419e597fb30c9e0903373bece7fcbdea202205255b18ed84070e645d43ef971547f60837396cae4906a6259379a46d03e48a0db2bc3b01efa9a3cf5da88e62d188d4f8c5db401635b8c6728bb9b65c616a31253c7905766be79b79877c027de7bbc6483030d78870e172986db48bcd887ffbcafbe2fc115e8ff6e983043c550da393874f37f20b57c872bde9e2aed200936e0c30653d562bf65c7fa35bf52c7e81444cea6b7ee81fd3c64b39c76a66bfa8e12306cfa10d1d62c75799f629161d24df94b7429be80cb45e2e0b6515b2dd8966cf17ef5a4179f7e4bbdfbad23d2a1cefb57cc026ef3679f9e148562b6d1649e26f3a73370f316c8daf6ad3a3f406f89ddd9ba80618306420260eddc58b1dfbae5353989da8396e5b7ef44f621dfda828aed859f43a353f1ef72ffd4b9f909bcc2b3365b8fefe4d374c9b74bde420371363bd2	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:42:44.323
hwj4bl3dmjht4qh3uqeqml54	user_admin_001	admin	update_comment_group	comment_group	md1b4mj0f6ztl29ksjkani72	615d383fb641b2b148141c46:ce1c56969f61739674959a161aa9635d:88df79a5fd6cd79cfc5957bca81b04176bf8ce7da491b86516fecfd84abbfc783dacfa4f9fe541efd822f0abce003ecfb9d450a5c7f156a1ff7aa5c786b6c31927cca35b76166be7bc9b9a748a2c53ce9157cb4adcbef8569ff7553c92395417da8d58dbf7ec883eb368ca93b996b919361a59a5deb8476abd1956353b82f34a272e16b0c2daa0187ce5622f4225ec2048a036d6de049b1bc8a92bef77e7086d6cccaf423c56b339040e652eb24a7224024ed042afb4eda3c7905077362d3c03	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:43:25.591
xae7pbgjyopt7pfd1oe0gmaw	user_admin_001	admin	create_comment_option	comment_option	x59xbal93acd84xoikhifcu3	64d42851c068438dac3788cd:a121cd40a6a339edefbaa70efcbdc758:c22132a6eb991a8dd80ae5620e9d721639cbc21e9ab37e1ea8fc723c2bfea0a461d240cde2cd1050ae9c64a9ca43fffd82309b01f84e271100c6f8e0160502bd14b1956577d2815320badce75e47d960f68506c55b6095fddc1b0f866ac023ff64772426f1f6c76f1a4e30df61d2bd00689c3db9281b64a1564870036b9dbd190b08b21f5a427a57169c0f957c25dd69718ad210469ca8793f149b14f911a2729841f8faf82b405646963e520dae69561f22651d72e63b35a0b2808fb59c69819545f1cfef8bd5a3c667d6aeee2e80fb04376b731ae154167c64e41fe360057f4f0ff12875623c5a335bda3458a9805e6f08c7403f5b539ee593aeb312d260afe53b515d61e7d3f049322cab782247bc718ea9a410be2855e2481a8efcc1e9bab706f4c439356f082199d4669b135d0de7dee963e511324537234fe5b2451278b5e6a955c29bfb46e510486bdd319eadf66c2f91573713584b1cfa1d0c525c63ed0711803f60301dc5f42d49f953177f2802807148048bc668f485363ba1bec2076e57c0870f443d084a73dbc568e5ab5a68e1a7af7e37c44544276183f3de6e616deef4b0ac73c1df8959d854dab11fb4012fa4cfac32d64e26ac5e23121a	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:43:42.932
gy4orerjgfvwzfedah13ngv8	user_admin_001	admin	create_comment_option	comment_option	h77f394372uiiqeppzsriyfi	a042b8e4d8e514e7cbc8a18a:468dff2ff905fb9c741028a1f2e7e923:4dac9dc2e391ee7ec25a1db803042b1b78a4c9baad4c7b51daf751204cc7116def5e64b65b166e8409bf66783f59e127fa0b3f7dfa45489c314b88a3cd57d6addd982df8e0508065a32cbf70709e8aacd08753afd1690df91286fa2f7b60c4ae9dfd31b2c15d673ca5dfea29d78379d44b1dc05000fcebf7fb08542c7bbc63cdf06bb70cba2aadde39a603f1a03d70a29502c208d1b35ae717568cff113d73a95f1f52fae56b5df6d28fef00377411e7cfd5bd782b9018d0efe6fd917a45840648500fd1e4fd3312e08f4ebfd4bf5a1d58b3df6164f42f77e8ccebf85e2c0e44442d910b887b6b882b709c490550d44e3cac3a0a85d0983f3be38e38e5247a3d4532b9ec084878f09edeab47034114b4b2fafa33249f1393b7f5425be1386b2f4a6b19355b9adea7e5550fd12100728863c9750c53ac9804c19650915b3da22d8e07d37acd443bf83a2eed12a7fce29cefb10f857c42f38705d586b91f631be926f187dedbc3660e13e5f666df16b55c2e0339e8ebf441ac9084c505483e3c4c2912417e65b3bc575207e6788a4e10f94e499ac87c24523551af50d7b5d9ff76af075c256af7751c8042ebe6ad50459a9a41ddc17ce04c124d640f6775dd4596bbad42cd	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:43:55.795
o4au4yc8p0nf65ed9t5pm2n7	user_admin_001	admin	create_comment_option	comment_option	owrb9f5mnp5gp77sjgu6d18t	e2c6b75d3e6fab32cdebacd8:5bc6b2b51c6cf8eb3a9124fef393b8ec:9b9999e0ebf13f1c7cad41eb40a8e49dfd6e75b2d73ef0fc808f24faef57e24e0eb14dad89360be86be057756fcb10aa12ed009f4b483bb2d32011b7c68007e31c9b745a0bbe27a9f6d638446507ec1286ab76be288e8f0559a7cbd6613011a97cf4db07973ef6b63aa6d59a701ba9cd4a9b4859751fddf4a6ed09daef17e738d29ad4b45fe84a3114585fcc37eb76187ee6bf1b20f187da82f89206e8c7ca599ed04ad57c7008d0ebbb98515618e3504faac33620736b1c28f9f75553cbabe55abe37db871e1cbc1b9ee699bd5339db727127980045faf3142ea328f678efe889a9193376ee4d36822ddfdf4e4c3ef11f8b0c02182008876caed3481bd9b532182129f3fbc9801b62fa76011e1e04c5a7360c804a4a5a22d218280bbbb6b09a600ed5ee97377f70c35bb01288972ce0deb839010c08367bcf7ef0f57803f33ffd20a77fe7b4ded5c35844de0747da6977fb65ba2c17ffba7680584236e0252e2bfd9b0afcb451b49fc724f3abf67b1ed46ddf260e7f9cd95cb47487c6584ffb60f7d102f3a9c61026d40096f6df7055a46f6e82683f7f8f6397be3e1509522e5bb13754c31b6cc958bf47ee5f3a8a4b389be51add02fec8f005911db41821467903745f1239d73cc66f142c157b9529fcf514	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:44:07.894
vq7hnc2xj26dfui23dtufeyk	user_admin_001	admin	update_comment_group	comment_group	yb8henvd4z79t4xljb1ciu3q	c2827228fb4593ff9b249de5:2aabe0cab7f640a6963bd67c2c9a9642:3dd097f16f19460899f96e940dd05e2ed6d912c73558a917fc81fdfef486026f32ea968dfb587ef5d68ba316ef3da7549e3a282bedddab67315a4dbb2693f057e65360a9660cb6715e6d28baa1834ad491cb90624dccc32def248521fde60e87c6b166a3151cbb09baa8f71be79aa97cd407c6b681f51e85db4bcdf2da85611010b540b617d9424360751511ee0655650ecd48227c8fd87d1a6cc88e280aa912c918eea2fd4e260537edff072975230d577f6c26d04077e55b6d362fa5e94b679762dc7ca7b67ccdb4efde3110d6a579977bc7f5b1ac0b346645d636	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:44:54.835
ml0hblh3a29m0f3elcn6g7kn	user_admin_001	admin	create_comment_option	comment_option	s3m64yzzb2xbsic8u9e6md4h	9f748dd38847371108c14953:a0964b19d8b6244a1ad1a13118a1ce54:bc4954ae86a9adece93ecaf4db392bbf8ed2e0eb41381f21354bac4c1231e8688a18c9c48a688b9eba143a215b387cb34696969bee7c335f8466e2d1f1ef45bce4e1572f9ae69326779db89eda771d40984925e3aa307d0027dd86be5018b5af911c19d90a82d4afdf0577eec9bb79cd57f131af35f8e5171fa3640726936e782e9dd0ddbb0033b1970186db63f6e2ee39c2c6cd199a9a066296520fdc59b66885446435bacf6c1b29c97bcca9d134ec131b3e7254928478d972f64b8abb99332e426e9d23fbe1701fb9466dcdfe6fee5f6cf63d7e10cd2a58846946d101b160bc2738e7715d3b1c4fc29662ee66196b4b9c0d3055c0379729fdd74147ee0b8bc2d845c7f3837366183312248a2422887131203dd637567c59709ad4e63582f94f6be9e68d6e8411b4b9c819dd55d1b5b4fe54319fd2ddfd910b9ec69fd325e39189a79d17f110e9e996e3304b6602faf8c7eec4abd9aab37c1ec32bd50e190c32f96d7bee8795afdcba08f363bd16c557adc1bb65669267cb849faa9957439722bc7845acd2e42b578769023a48759ef0db9c78ca44b205968b4179f4b6b991c246a6dbe7561c8bf08732fe09d22f1461b4aa109a6e7aa7a5c6ec36cd8709dfdd4a2c66bae2f3ba13e08ea5a9aeccf59d72ba28f18cb2e3f788dab4d5e2a908fd8c40dddca94e3c7ffd78bda689e401e4630c26e5a459ea21eda5481b7fcd74d2822dbeeae4613e7d21f1a4f9f9e369013e62aa214bf3a9757a60a61d27ac9337b033a5c19957883e9e203c993b408971acabae27bf843f840288a7c18c063bda0f1ff5251ef47215a81aea3f804173d81c9f	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:45:10.174
m7ulbmokdtqeehtokpta9zso	user_admin_001	admin	create_comment_option	comment_option	upb1bj9odzuy8djekhexwo6e	ae02efdb1751d1fd0d69b75d:4ccea94eded3688af49de609afa13524:c55c35e2955706fe586abbfa72dfbc5469c1cea78e30139d1e958b55c3d1a01efbfb57991a05c07e06a427f320972d3d9c01ab41fdc05dc79a6e84dc480e8963846ea3f555f994a1d530763f1bbee74d4000f2ff4fc975935157b7d1c89030120d492c8ffeac84683e16950cad5673808ef1f49e4aacd067a656b4ddf2cc5bb5b371d8314f2cb2894f11933a5ef9f8af4e4e18aedb9801b5a5fb73fee1c07e0b83f899abf6baaefa2246792b93b9525a7378e5d2cb0f3429b8bda1dc6916a32c46cde304decd14d528a6ae04c11c26e3017133656071523c47c71d3b545034789db137a2f8e2edebb83179cb5e392eaa3c79fbeb8f612c32ae1d60bebc5a5f3ddcb161b6294516ef44a71c84644587d05065616acdf872d61f0f8ef9cd203969273bc89c6143305d6ab260a4f4d186c3d2c57421ccb27537d5e57a1ff4f91296cee19f4b69646c997b714cbd82d6a8f2d341ebdd6fc889bac88c50e0e61dea804a21ee2d6c8de7631b777daf7315eb9f377f635f98965370c059e2bfdc85dc0b1a45fd9f50978ee3cf13bcde636dfddf405cc21a0b6cd29a359f5e0b0a6ce3ef74454cd03980ac6f253bee53699b5b60284a4a334f6fb4eaf7e622448c8a83ea5443ced49b587523f5138c4019c31591d5d111172e50a0e9a0fd46d6ce09676ba6e6160c7666a6c1d093a978c43346cd90bbad16e7391077f1cbc75756ddc05ef5ab28f2e544c0545e3d7e3e7461878809b8	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:45:23.815
kr4i18e9pg7dlpejfi4ms8ae	user_admin_001	admin	create_comment_option	comment_option	kasbyby331o69w6kusa1u88f	15dabfa0e9411391068ccc04:99a6ec747172fe567d8e7871598c0746:ef1356a831923e95b0e7ca9f276c6af66de0a764ac603345a54371b83ece32efab3c85403badf73365405ca9206fdcd7e2583ab35441ae6cfa56e2638204617b48630b7055279ad32368c62bd9b4992f294d4fc75f5e90a7494c999a62a30f507d679b463d2775677a66e2b7540fb7f6937edee48c0ad1f1dd40bc76b177d9aafbbd44849d15bdbc6659544bc3d2520a8ee5cc05cfb1364a4328c1e5822ef5b2a5824f45fe5be7085d90afb5a830574e6ee0c2e7445425ac65cf6595c2243c32bb1d80f2f9881c5eaed9b1be12533d1816575d41d29d8b817fa4140d4ae764e16b04738407be3af9c95e896ad3438ee7bee936a85bd2074c4ab698e8204c6064045fdc69662fe587f4215d33fbc14ed23866b5e1b8d264f7716052379d74a9fcf076d5fb612a3cff74a0a222b1a435b4fc16ebdc0bad63ada68ade9c4e43c3d5863adf781424cff1ddd566d1a8aebd7c7259a3671d33dbdcd0bd8a5d27e6d0ebf58213903fe2834d6f2239667d572c1052346c4fc6c552d011d75598d10e4d87750b539287747e0aceb473ea5bcfcdcaad5288dcea58042fcd5483978292f4e630932a3ddc1abbca0b7f6292a0125cf03ec37c9e76ce2f50566f6a00f9d2554f5f7bd6af105c5f673b	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:45:35.345
tfueakq26uv10sb0tjeys9m4	user_admin_001	admin	update_subject_comment_format	subject	DT	d3064892ea67e86bc7356dee:afec19ec69db056dff73817dd984267b:05af7af12ebebb2e121ff44b69ce5f513d5389c07ca79cd4dda11753d01ad9d47ba58c9dce688bab7fbd76d53ddaaa523e922b32f877b0ae756a8c293e2b07e93701a49dfc485cd2146b18b958acc8f2e776ae12c3cb67bd5b5747bb8a5f118f22ddd2848913f23cd5bdbb9d8eb98ffba7515994c90fce4eb804e0	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:47:03.522
th4l21n9i6pjn0f5ksjpqt6y	user_admin_001	admin	create_comment_group	comment_group	xdtpcx4ay2u3o0mm9fuuso1l	34452192333d3fccb5ccbc4b:36ac985bcad8236fddaed86941325b45:2e08c6be683ba1b645da4e24340198645fca31cffa2a7082a7bb719abfcae467adbce3f488e6b305ad07997b4ba420c8620aa9934c57b8d1109d630c796b0f1721fc9f0578f01a12ea8510d37cbe45b63c1babc1687bd74612ce3a4a6165efe5748c9e93c8de22c67b122ffc82f4	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:47:51.085
m4ags7umo97wke7xsuzr98fr	user_admin_001	admin	reorder_comment_groups	comment_group	\N	47450eff372bde8f46e165b5:ef3c28a012b0081684683b2ac5cbfb8a:84d1a7856b9517fd3e58a8ba178a5c858168bc69756abb197815084545d20f4aef06064cd67d7b1bc09badd63ef52145b93270fbd2d24e0bcc12ccf61549866f90a6f0fbabe54f28722109b964252b14b4c51e7abe413181443b970c499a722998447d112b053e9e9acfa070745dc1851c5805e9b874e7375eb36232e3915adb30d3276f1c8429fba939f415cc6bb0c44a9ed6be249cad41ec7e485fc0fb62dd890740111cad2a83217c154f15286dcb389a3850dafd1e4dae61e2eedfde57b119e5579cbe4fcf4357d0e8ac88bd11	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:47:54.933
jcyzxp8199u00vq8wnix82fm	user_admin_001	admin	create_comment_option	comment_option	bymqwmprlgflgl5ps40bjwdm	7f94f23a1a36ab32b284e12d:0d603117cdcaaa34f1b34ff5da4e2daf:02a0070f9f6be5cb4faf8f97dd159e82dace3c24e201c9a4e962d11f83355c31a1d9d87e71843b53b0515346e50e4bc9a4126c6481a1b896b4a31541662a512fb33e10797326323cc7ca047deb2fd254562b4c6d7daa19ab99ccb6d341457c9387bca9bc2fbb4e2e61eba733e50f8be684b7600ccb63bbb10442246bf21010fe537f45626f5d74328c362d20a86c1d07a0e2c1eee5e2858b5ed3a1585d8b384d6957c4e5c10e8409b60040983af17d0757becad3b4e5ea818f3cb1f2427a677d3c14a26d3d65dd6dbb6a2bc2b6b28e4e3a080c511931f539e499b3fba557595b22e4a1ffa02ab10ddd23b8ef748543d4c000b97eded53778a29b51e5119eb34a3c3623a74477e35fd342fbcdf7a514819870cefdb26a0fe47328b1dd90bf346c797b01311365164a9a3eb91fc8f38d8cf9328b57187a54b0dcb3418ef137b519721b69c1fa361d0e5c007bcd8d8dace6751d5839f96587a1f22da621dea5f38853431eb7ba5f82d807d330100d630137f20481b54e5b9600a8f2dd8df8d16aca1fa1224d1f7db3a6a9bda630f08e19759c757495062cbdcb60b4490372195525c6170fb1d8bbca11e9bbe07cea3bf42388064e339dbd278c898a3dd9652e4faa93eca96e0df7d1d766c99c4c76d6d4f9493172a6af31064dcbace3f21b728d2f45cc2de407533c6304d39c4ab3fa333a1c79eaf5708d8fbb73ddc3144ffd278bd65d96e2ce5b0face1a0f0260f5ba2d7c55234792f779a038877fc08185ac3fd72fd0ad176db07285f2f966e8f090e259a42fd2887b3144866db5b8e4b9bdc26e2d595e98a43ebd8aa14c0edd51d2b9a27	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:55:43.816
rqfpini5y7qifpr9wv7ay8k4	user_admin_001	admin	create_comment_option	comment_option	ll88auca6utxs8je6ivdds4k	4d7c4cb1bc3e070ffa5a2bb3:4104ddb38bbe408f763327c7b26cb1a8:170cabaff98162a15835fb224fbad4c4b4b4464ece5ab0d322b49e7bb9e29cfabc1ad8b464b9f4f44e754fb06a41b4b77689067376d2b72d3baa3527078d7c2ca7fdb85bcfecbc4f072d51eace271347b1b83e034a734e2a470b4f784c20939105e16831fed6a63ef1b234dd0593673f826f7e074bd57cc8fd73670dd10944ca3f1c732f59b5d99c00a3639c365b32c79ccb2c9f003b60c2048cdc6f7effea0114ef2cd6ace5ad91d7dddbf91e62fa2bc192ad8a99e852aa726e2ef33a1efe5095e1f9095b04508bc30e04f8a2acd848d05b1d694f6b4a13f77699ab7a524442a3f58b553bbb6af1292afd942579f9c6a9ba66eb7d9c896a29f6470676897d24cd03448081a6929ae10968a954c3392dbd132ece4ca70725f607f89a6b29d110dad073d07fcb0f3d94152e6491c8bfdc11d463cfb2f872b835d948aeb0b9e51efbd25a783cf6e35f70dd9f502597f95ffef47539164d4b4e0240a33ff98fc44f598a03ad1aee7638ac1dac1fa02d42fa2dc0fb31bd58e6e9f374de67e2d99a09a716e62e8eeaec4127412350e53f182e18a0d49bb0a7f29fbd40c3ea3b9323cf626f7dc491334aa7f88ad7432bd69c341f09be4ed7e380f3519214cb9fd3fdfdcad2b5c79b302b4aa53ae48b785b1a30d5effd610c23b8056eeff131148f73bd9edf581c59c71e89447ae5f33243	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:55:57.918
q4catsf01a9ngt4d067n78m0	user_admin_001	admin	update_common_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	43f42cc079ab1ea0dfb32796:65e719916ef4282a190b58ca59b6ae0e:2d944ec547debca27cb46e3d9ddcc622f75db00fa5493192d3d31372f0f3fcbfc940300905b6b3d91054465e92a4c4658132c27266336217cd299ad132536a32926ddc5304734cbdea562c0bd1fe57327d51e98e09e7cdbd8e5180ed5191193cdfbc04ec923ae2415a789aee8b64d9ad69c9df61512208ed66c307f0e208c84d04ed9dd23d4b	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 07:56:08.949
m8nt6bd1o2k8ryus1mn9djmo	user_admin_001	admin	create_comment_option	comment_option	kzzm8ki6qgg2n4f50jadnmst	86a11309c40428f16b3994d8:9f68ff9a156f4a8c079a3ff0288657bb:8a7d4cb4386409f84d4061c19dad6b77c047372a4ea47b01cb1010b084f7bf215a7f5bacada9aef2df89f3ca4f40a88c51aafac464783be6b22680f57290baa291849911c59266d37d030fb5543ad805cad136ea29b2f7c8056b5ac28f66041f8928bf894bf642f8e49efc518873c9335ba4389be9f6d5755b70e48e8afd33b8ef6f12cf7a78ca179828c95a819e37c4133d415ba8984a538f3eea1b4734bd95ae3c87e0425792d7954af521a6902d04c3951645adea64193fbbf659c8e91f6bdfbf3384b62737026b5c0cfc56bd69c93a0ad6a11981a7670883cd217183172ea60f478521e820c6f83df2ccacc221cd9480921e22d641295e5235910ced9ddbdfd691e6df6e94438ab5ddb9bc56027a98bc338b4abae79244b65de3db0b2dc1c2815b13d64437a8baf51e7135e71a336c2b00a4fc9b55d9a3ebc6b4f5a3dbc8bb91d44818f7c0d7cdd73f410e248b5854ffe8ccde870fe76a59c01e769440357540a9cf736774cc0b83e011d6c95d26a49ab19a19430e0f7cd11fbbc01b62e10a40e8c2af4a27898eb7151ad2ffa3804e4424c65f63f9e34ae81dc71b80ea94eb00bc1d47095c840b4068683df94463c122ca2f3424c0468cb23f0345a7526ff615ee2c483f324aa371efd6019e9a04047d824b813f97df6fd5ab3f3986ec4ea04d4d1a8ad34810a3a6d4082c5e236ab851683c0d099851093910a9966134aff8cdfd9e03a18f909bb5f5a0e9d59a3e	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:56:20.529
za17yptrrfsx74moqqv989jz	user_admin_001	admin	create_common_comment_option	common_comment_option	tws0wzviep9dhs872nqidd89	dea66f70cc39960f663c9b5d:5ebada68ecbba12612ce0d56377237bd:fc69a3d71d4bee69d59e478160b27269a7463592749d3f30258ced6e58bd09366032ef2ba0c5134c376a324fb01ddb5fc0069288e996a0f113192eccab8fc0852f564263160ffeae7d9d419a0a7c9ad74995578cc7b2f38103def055aa4510f47e1cb3eff0b89b00eff7d6671a84a21e1f097dfcb11f87bf9ba25a20aecddd2c7cdf97eb368619900a321df4db0941b734cd7edf04fde8839631636b6638cf75769010af213732e6a821f69d92f3c4e51226486fc9e00e3a80d6e2c09846c1167a08c879776c59d3faaca861bfdcea0522c5c0b7fecca7e9c884a683612f8b9c5b89443cc4bd4c4437ea3aa653d6869444062699da7e8fde7d8cc3510dccb4e4b0f9aaa911b1d788d502c6eac73bf556da9cd3c0649ac3585ab44205d45d198d3ada1192411b2ce2d02d2003c39b920cb74adb4b91446c19e839f90444d68ad7694831fd0ef9b03090d64b5dd4feba81a4825ecfbe29ff38a22c617dd88b94633c8c3b25dbcd41bfd42b56b01e319bf737758c3e667a51d5e8fe45ed9d0fcf248eae745007ee199b7a6192d93a25b726db3c230c5a5cd102abd2180042d9aeea418d1344bf4010f02a491aee72d3d8bc3e8336ed47606c191e225d76	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:58:42.489
ita7vyi08m9twyj9bemix6jn	user_admin_001	admin	create_common_comment_option	common_comment_option	ykihn6of36l5fylrwixku2ty	0e207d0f65bee5d74cd8a85a:9daea8bdc3351570825c0c7bb91be012:1a150f6786497e8f277ed84c256993e1c3f9136a6fd15cace4e20e7f2e2cc0435fd4b304c5638c9ae41ea2f0bfdc1f16496cce599be12774fe59dd22a0e50ae11bdccc7f2a25f1db760059082326f21e4d8e3ba76075d4de2819fb498f82dd3413cfb0826b50feaed5ca8d3ee69eeddd5a99bbcef2dc9f178a667b24025aac56c8ea67b19554aa4656658864114f401b9cf29017c1b5886621a7b3c927d7d24254eadb1ec919a6a871ccaedf1b48bafcef3d87915b2bd48d859099b6426aefb2205d93dcef06ea892a51b64df0e646c93eec280d611bfcf7d3fcb4b2b275605f4cd4e4c82cf11e1af69a7d7d42345c7877cfc058381ff7c86d70caee7ee2269fcc033836d1d837747df847014591e0c53641a931ec2cd5bcb573cc2f844ce263288fbe3a5305dd4e6f3d26cd3e2411cde9eebfd30cdb46ad0b6e20124358237d3e3ec27402f9d1a96e02b02dd1b030b3179f91eae8838bd6e9accb68c4f5f02750954fdffc1c8a1df602fdb9c43b619516e70475e03f7715fb3bae7f0249e7374de2ffc4a8fa3eedee0d12fd4b41e3	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:58:55.327
beetwdaygdc2te8cbr8khn34	user_admin_001	admin	create_common_comment_option	common_comment_option	ghklrglxe1szmginmmq4mycf	a3081b1543470e7e92012ee0:51f4959e19068f0746c1694b395d1401:ef99c39239c62cfb8a69deaf96fc3d9de057a44a193a12b3cb0b2bc58096d467f597bfad73b052e40d288a983446bee79750224e83010483529e021f8c9b00231bc6285bb33ece5c840bfe4583f075ce9df73f4ec909357f850feceb2868a2dd2dfc23d226bb1be387e19bea883b7ad22418ea67568d07145b84f637754caccdd8b596ff6a585b1d301660385bba0dbe258113eca5c37a88b23586e13a4868644a80d7974b48c08b04ac54f1a7d7ba055714650e767b5065f1b7ccaa516787c5944d49791d0e0e660b30bba3f3259726a02be88a7dedf8d81ced128d6d4bb4a019de457c11f39939efa362bb6ac63790853ae15d7b8a39cf80278cfac1c702ceda7ca04e92d1d314b340ba21d443f56d906b32b6fdab17e8390b19d9ec618d6227c73167d6bf8c854a0334067d56fb26c808297843ccea10a87df73cdfb42ea4fad4fac5575a76a564709ac4b704c38d5e798265dc26935bf0b6b126a429dd757354c50455be267a00ef51a005ada17abe9660c4ddf87364b917e036411b6aeca95a45f14a3b9e188c64ad6322cb20f9f915ec1808ee56d152eaa0b5cbf1d4565804e4f5a22945cd61e842ec348ca932caf76873bdcec4	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 14:59:08.74
kpimduacd97lfvod976hz8je	user_admin_001	admin	create_common_comment_option	common_comment_option	up91qk5epq44y6x6apw3x7p5	4989396bb1b5478699fb2c3f:02820657a4c93d60f7e2146c00b7cf08:9bfe3c5e608e4133f2f3efbd25c5f20f3f68256d8dead8ec55038f9fd167f1e1818c762c60cb02ac2696617635cf32ab1f26bf46737a1d27d8b01cb0056923934c35fb3bdfeb4292087c37552125594f43f917871f9b86f7360532bb128f6aab6b90200fc74f6c15941d0046e089457fe00a24bd7548bda17abd3cbd08b9b81727ef7a02966da5ed85d9fca15c6ba73fecdb8f9ed5338b7e255ad64c6e794d2e5f924cc55b7ab98b1fe6f3f0e714fc5a8f90065baa19b0864ec3fe07959d97716d70f81d1521df6817ada9653029591088b66f2515f04e7c9bcdf90353fbb78dbc83f9ef8228d3a964260f8bfce67f8de3a1c6f7e13a73d349176677f96ba7eda65d7fef9ab14c4bee51f2d17753354eb0c456eca9bcaf1a007efbd591d9abded7602dc3a1dee58186fdfd687e4cec4329fd546e001cc917fa591e41aef7046714c602fa56a1b607f3f9fb537c58283a0f7986a11e721c92b79915ca2c437ca0dae035ce988e2c163c77b1d02d35103510d61f50d813f8f6ba070aa49256e2a9a24f29afe3fcd19484c6530b50ffce91502404e73d45e00e1b23e96dc093a7328bdc9fb43fd7fa42404fb3756ed8b0104176f30c68b56b45e8f48cca58	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 15:09:58.181
ghucv1vjmbmcvneoocwu7dtu	user_admin_001	admin	create_common_comment_option	common_comment_option	n5nkfa0ysm5wc4kwwlco5hm6	eaecda3977482fc275f798c7:fae1f9e5fcda224b2aa91946b60ce7ed:9ee3b5f2a8cd05e61b48383157038ebfd7971f4cf189656de8a62cfe8939fbbc998c45e4fb7d3659b742391593a7b2409b8dc2c2b587bf1c5e91ae5843d98af37b89da485de9fb34553d47b03dae3220fbda1c7be37945890b03a8e715ecf6047ec544b56882d8d93a736b04ac31203263ada2ba7bfc9c5afc657de7ba4f9874010e262a3438830e30898721273838351e820cf2a120ee11c05715d3c7076e88e27092c59927bf794193d8f4ef8bb325acd2b66c0426ddc8ab43078767403852a21939d886b310e0a8efa40e99b981ed7cf87d02d37352c8d1634070e56545ba1444d20aa47d9153036d810d67f06186145257f4e0de449b5c2bd7fe27c886df8d152e052afdf1e15cef0ed30c30b48affd8fdcc7fd64380524cd75b94d6c437c2a64384bd74041a39dfe9043b10b1c05ca47334b79754b48068c8dd889469b256f2151aa6ca3ed531b1e410c7a1fd9dc729c98635be85e78d18e9abbaffd05769abc94f597463306ece36c4e70306b59a25d30c95376e963038d0f449ddbf62720a868a52da78da5c8d9cf9820634c11ac9f1655727425473b2e3223111f3dcebcee5a7b8610e	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 15:10:21.491
fhco55bq1dj4wp3pu5yi8z6g	user_admin_001	admin	update_common_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	01a89dcb9641b6da5b59c096:5b3647f8ad892bba1719a152ee9f98cc:60e51e317b41bfc2ec552ff3eac7c87440a9bb71a4cf8f1b618c7c5fa30582f3cacfead4aac66970c932c7ab268ee7e86973c720ea3c838ede8a45c5abed4f531101681b9eda92b8c948cce5d6aca3d0df99b3b899d679261fc0dc07da50e94652bce8f4ac5d0a92b42f75f5c6a15168ec7b7033eb4ad51c581b53dc83c808e18174e393	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 07:56:10.895
xtl7diq5kre9lsrq54c97g47	user_admin_001	admin	update_common_comment_option	common_comment_option	n5nkfa0ysm5wc4kwwlco5hm6	c1d26493fab2361ed7cab38c:e8e133e425b1664af78c91b3f1d5d321:08a68d08886b77ebdb8307cc4e6a31e7e5219d2ab1a730adde42436a3be8b3ac0281ac08e9b8f0bc3db44d5f206ad5e01e2a22f8a7fb9a6a3bb003af719f4a5d1190866efd27290cbedd1252c6cd74cc2e1d88cc8e9c749ed2e5362e60371929e6a325934814c128f0e5db87be0f1f18abc329befc7fad432c396abe470023925c027b927b093bb122062955b6ce248392106e84bf14fdb5b535a03cdf25684af93d99ff04482a10078dd054b48196b2290c07e922fa23a58e98e05535e656c42128c0b884c5fb717677b8f02cabd83ee6285acd080af85c2884d96e1c4c67a5a4253bbcd31f99eaca6e4528b2fecfeb0f7b1591aee4ec5993bbb7e027b714bd075e0e3aaf0afacac023cfe70aef512338b0b741c3729ca980de70361a7ba26841e9cbc211c03ee31b0ca1af8cb332edad974913a24e074b90301677155d1159da70017dedaa0c774212473e2b1356b69013943aca7c8e4b5ba26895809a17edb9dc698dac903318635c1d1d2b04cd25ed89bf5af0153695a5920ff0bd2a82bfb7dfbbc2959b38d17161c0edf13c08afb1bc0594a143dfed3adb861028289ee41557fed17ae6106339229d101c694ceb6c4d682ef2bdec8d5d1cb5f86bfd6b668a9d436489f0807aad8b84601f8b42244e33da44818dea8094f0bf7876e342c88058c4c8663c8c23339b3cef0e1b95b99a26bc0e24c4689ffd6751e69f4ea9624fb3f31aa9e0528a05fbeab984e0c4e746433985c7f09f3d39e6b275982e245a95f202490d5f37d7ea378b5a676d29216c4c7ab1f8e23a3b90f10bcda30b7166fe7cbcfc4ccd0c9637bebf378d67189dd22da707570244224582efee61a775a30cb0d41ed7303268f0e257c9f4dcba2b98e049ed90ee09c83964752afdc11fd26f1ea9c1a4590a586c2226cacdb55e0196dd5fad3744e2933a79bc65f43429a8e3a06925b1b57269a4b1e331828e85e84c193f1d50e0298b59bf7f74f08478ade18e5d0006af4941dc46ffde5fb6b6be2d2400ed92dd6b609208d0d11db453e290ce8010affc9607b609c9e0730cd14fe79fc0fc1829324f9e88ac7ff85a3904abce9f3d	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 15:10:44.074
sve8wt3z4n5y31dc6l2ti6xg	user_admin_001	admin	create_common_comment_option	common_comment_option	ma5e59t3bpicl9secnlklsda	26c7cece9f92f2bcd77fb1fa:79bc0f629e7bdc9cd1549dbeaea4a778:2d67b73e848694765e90367133625f0d15130fb0926d74167e391b0229e6f614990e1ee6ca8b28bc6a471863d8304e459a33affaf36cc2e6ae3e429533bde184cff6c79cf6c50027ea9dc336b7a9a1235a9f9e8b8eea443f9a2c0726edde6a624145d059b494e1ac10e9d96fd08503dc78e72e896bc75c0b1b4a39937a19ee7e94c50b224d9b440e06ce6c757597692aa985fe3bd50d23620f0ed4c8758c3196232d8d935bd0d6f47f7880b839bd2ac69087e39aebf48ada6f9da640abc2bb9aee863ebd01ad5bf8c4eba2e8c8f216a7131a6ab7c6b8da5232a9b3550ff0cd18e28e6e59bd873d2a70ee1643f16379d971dfeb492df09c0bc95984547f1e45540cc824fc404de7d8869ffbd183e491c16e3bf92c940ce97cf57091abf18ed107a36b1b576b74042c5b632d542e5d41d2c8fdf2803049d6c71e96569d5b5e4b77809a75610048cf54c29ad4ea76a261de452e27780936c4dbe9410c2cd9ec3f1d4a205a1a741e6e642d68ee1a0a171209119067640bf658452eb2745da499c624066b140ec73976c7df02912c0dd06f43dd32df3579359db4bfd3ac56df735570d1aef23ae751fc74	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 15:11:04.617
umi6x70cfd4nea4c3ldl284u	user_admin_001	admin	create_common_comment_option	common_comment_option	tumuxai35otinyej62mb8ceg	909c7ca0029947b91941121c:0efb49d6d1f3ab28a826ce8c2be367ef:3450be3ec207646ac0950eb193c3d4cc0ee834fd229f65af49c1870af7b802c34c079ddc972dc7b05092aab0fc34b94d006a80751b8cb409d453b02222602ba8b3b32f72704d0c254613d5548383f7926f0fc4992b1ca350f6d55a7fc60ceba5994ffcf13d359b7d2870c0a6b9c37e1cdc3be022dfcfad78477242c49b577554b7c1197f1cf0cf40c5b728267f5bb1482c57616ec45a773be1f1aff409e6d94cf80ddfb749405329c4b0ac91e342bb2ea219b1b7927a2d752f43053502a44a53bde1474eba6614ef6a8c418a8a601b592a81929c1e93bc012ef0e59ce9a632fbc4e7402681a7732ce092c00f890708512fda6cc8e7a10c1286e5b3daa8e7d3b2079a5c71e5411b514b99d6558cd8c604d72b2fc6d0557a960c3f26a351630de902218cfb5cd09c4c0a8b4305e8cf0dbeded62ff98cdc46e84be85a1679b47c36569970f22b01ad9ca4d4ad747944ad321f10ec9572e19709460dc0980119b571a819ab49ca5fa157cef6dd1be4f8c0d2574a1689042cbe0983fa3f068ba6abf5ac64bc68d1c24b6c9f3b89067e582ba607fcaabd430e8e45cd15c5006138e11a20898d22b6616ee058e216527ed7684b592f0ece335ee4894fb365e8ba98e214fd1f312382	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 15:12:03.638
c0ryf9o0tnjco4x3u0uq8l54	user_admin_001	admin	create_common_comment_option	common_comment_option	b88c0a8lneuld4wum8roaxse	6a42791e9e46cae8d765809b:a219276a242ccf312e5dc0f118b55851:d52b6b3670fa7c9c0a5e5eea65dac921fb4488b2fad2dc8d61860e604debfbb1b4721c9dcf09234fc7da53a00d1b5519c496dc9a681c3b6aee5fc5ef9f704d65767306580c5c76f491789d1ab77069a0623ca83cc94a058696696555707b6b846b3fc79d7168af487fa7e76f26cfeb3f366d5b41b4bf11691b7d7c5b9529ec8bea04ecf59655802de0edb122a5658d6f735cd9f41d85510ae9b07645f41c09778229994b22b6e1dbd69a3636fc3c666f6dff38b68be566ffa736a095530f343451c54498f30309cc5916baca66cd2d17755f6a2a8461fd58e24e36306f3a099b9df148976ae7d76ab29d573c8d5c241823df63d72191b422b067d5727074779c6e761d747ff90b1e911071d6abd65594d9048cac5e9ae143f4b7cdc5c5d2a934f2b829a15f54e622de7a8480b5855819cdc1be9e4413ff4b9e21d46dced91795f001966080b7fef378fe9152fffb52848a13f966df981498098c65fe9662a9d052cf5a84847b11b152918e8d6ceab0b51b9209a6876381f05cbaa9d0ef1c65e1b21d9d7f77839ae04418a8a0f436dd815323ffdf09b110555f62c23d112b4625663b456ff0d522	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 15:12:23.921
wbamao27jtopzds32f58uh76	user_admin_001	admin	create_common_comment_option	common_comment_option	w01iob6b0o2sglgzl4ipepyy	f2cd1dc382097efca555f11e:570ff81427833eb47c49fd85bce4800a:ab93b7c630fc336d0c29502d4c9ce7e777256746bc023eea10595c8265708c7485450670e73837c25a4e681b49e0582f2e554fa8d968cf308d36d7f5075d3283dc5e988fa8ffc5a1b969a47ea0779bc2f49527697b41e9bcb670b37f7aa99458bec71502e02cc39483892e2d35607596d1dfc4d2cc07302e7ec0aa36980a31780af44862620fec751d25965b6745fe104e53de50d3a62d495e6a34acce445b161e150a8b48cac7150a66406aca6bf4ac1676ab4eecb5bc801ff79034d72ad313fe79108cebb833ed174933aafbcec82d82f41de96714c3d403f2f277357d1a38e08b701ae3c64d3738d8b878995eda3005a197177127d8d5474abbf1c8730910e20c2786e93de911058f6f3cd716c1ff6fe6ceef1300334ff542fb53d83c671a6488b8358e84a322f3613be452bcdfbae55cb811a0ab23c61ff5e198d8df3c4cb70e8199990a6225feb834c8127a5963dcf30c053afe754ceabe78b9a10485eeab05da0a5b69a37bf5632dacc88588267147c7ddea7ee7c2e5aff6b2667ae898a12c2ea0ea75fbcb6483d184c97256397ff629742db299883de090baeaeeb1710731f24e1ff3918beb9d06832f8dc8	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 15:12:37.354
zekdgr2hvlv8u1w8ms4f4ksi	user_admin_001	admin	delete_common_comment_group	common_comment_group	als7ui0tyuuos8j58jndt892	b581db06fb2c88f854d7af6b:97703b306ff831af3026831cd19ba35e:5e72349ecbaedadef7987822273d8408203fa6550f8897cdd6f3b1c6c90525299ef8cd9dd3719fd8b56cbc20d017ea297c2d7adca6e2c7aa8f3d4c61c43817d2	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 15:12:52.553
p36idlu0vvye4lbfanuin83z	user_admin_001	admin	update_common_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	2ae237c8e79c98b6e7a7bfd2:21b1a054ca578eac0aac82c6d9b49b8e:4c5f9bc44c8514b1f77e6e8b7541fbe75c639fefd2a54bbcc116e5d592a01559a82cf3b651d103a9cd12dfadb7a25ddfde0e50c859939315b0162abffe4748ab10d3ba6dd9b03dd9b6dcbaeb964c3934b06aac23d3810d6737c40b0ae4351fee6ad90149eeb5175d5a7603d76cdc669594a742e31fa4e9428db86bc319367f7666dbcc71492b	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 07:56:17.86
et3e74bkohxlgxofkus77nut	user_admin_001	admin	create_common_comment_option	common_comment_option	gug2j5s550fp1wq2pg0f8214	ef91149d94c3996cb54989a0:b393d7a700a2aa329dc6046a52888919:394171d7a1ee5067bbdc833eea2d50d0ff409660e5b48b560c198c36f2c0062791383313f1893241b081a2cd985e4c313ab6465a6c65af2f25d25522675453aec58c9a48f07e2806d22faf1202be109df5acfd1fc55ee6007a94e299b58b7baaad49718f9f63722df8877678c1282fb66e153f1c5f43df9e64b10db2a920c16437ed03c6936e0002e282bb3dad7cc5f2584b8f48e502cb1a6f079e251397c28a4d9f892aa36fc7291191ce14088478c3f7d54da4102310f417b78f28c0452d6848f27bebdff89a3a5db054c3b05dd81c20b8f37aeadfe1c4d719624681c6f8b116b00d40716830bf11ca5e2111ef2f71746971b1109887c9d83bad33eeb4f17c9329658301f7d461a0801f61f5410043027d8b3a0899be968b0661854b3b168b7a85d0b0b575974957c00ad99cbfac082a0592d683c8ed01f1b50c067d9065f3f4295dc7c3f5ba8128fb3155b3cae398ced89cf9e7678d012ef0b61300112fe12ce3927d40bc6ea9fb64bd9fce481ad73ade6a1da328ec4bfcc3bd2c7e49db37b1342dd9355763639034b5700a51adacbf9bad70db0e1bd4b556f8236b21a6bfd9663d131df88a2ffcf1cd006410faa57edf61bad883260be878239954ffa8d8a11624d7e18bf856b2b2029002f003940da78ca34505409edb8d93a67a6a08c4752dd3118393a337ebf708d0bb21ab3725bfa7a4e2c68b4f73de3b7a549f87205e6ea469d31b46abdbcc163278a28b209057f0	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 15:13:51.783
u1zpp7g3oxegqn7p5q9qg54u	user_admin_001	admin	create_common_comment_option	common_comment_option	dyzw23jb3hndcr8o6parcm6u	5c31449ef307dbdec742ccc7:cdd6bddc1b81f6abedf9ba762dbdc305:88154ed64cd5340e0a5ad9770bbe761771b7878feaa2069a48307219c21bd137ddf3f2dd5acafe37d95965be39580a511df7f4c595136cb0ac2758bc40cdc49823bb0991cf4975b04f7cd065522ed138c80406e24251eafdd6a3ba2cfc09f04a73d09ba74f08d29ad6a5a146ea0fef3b2042ed2a0e45ec6164cc483b05d3b5cf94b6a7a8b310c39652016821f508a9e1df97075628abb9f0d50d5d165554f55b4f6a2d2191fd2a8bfcdc8451641b3c191aff28aedee244b792d8e482d2caf28b0da074850133e094112966586ca6ae4ce7a5e365cbfea7ace81853d79594d03c407bccc07eca9ebbd1910c3b916104e383df0bf029c4aaa0f51cee54efec8ffd0e811e68b22deb66fbbe4945b5b5a796ddffba17eebf62c2b7d666faed8534de3cf2883ff99860663413a1474cc5d6c0c630979386dce3eaeb9b49e9f047fe6331bd4827bafc2dd64553fad667ea8cbffb20d1b078e82b6e275ddcc3c43f6d8de2957e336ad574689b61c53cfcd48bbad0c2413abf0a50ed000c5858a0783efcee5ea4b409f1d6871093df8139f2a21bf2e3d2fc8f3d7957d057382958298242f2a9612e32ae50b7c96302baaad316321fb10a9128457c87e2e502e53b029bfe9b1f29f4a24ce3ccc985409f4ad0aed90915f6a3dfa2513610cd056cb764fddcf03999c7313a98ed76b755d6772af727f34e588c16177517b3a9d88f1a2b0c7a01bf0d4f6014b75b20039c25316b499b670c3a6b	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 15:14:08.816
ukquzjqv3hi9g5h0tmg771y7	user_admin_001	admin	create_common_comment_option	common_comment_option	bqp96t6yfdbs2lzx6cdnf2u7	8e50604613d7e03705982667:c1b2472bbf64eaa0f63744824d24db08:ba98caf309135283a55b9c98c60f548f1b5dfff9800d4dd254aed4061b29ca0ebb6ed8084897223812248a6a79c03cdb11f146ebf969c05f65ffc7515c990d0867a6dec83fd8de0a05f8af6592cd4c528cb83878f2459fee1b1517b42013e6b4904cff1a99481575d772af697c709d0f5b671eb311404620c3c60ee47d20cc1d0cb96cb607d9643a8d657e2c9eb6ed1f7812558dacec7aef061522474d841b7a47ddeee0bd8d873265ef930c1026e0e77dd07a84f5d894555efe7aa1a971fa4548d215fc016ae4efe7c3bbe675082ec8cfd9ad167063fc349485ca3cfd9d349c84e7196658a02f66beb70ae0334a962007635ef244dc2b1d8ae374a1579b0f435a4d3104e39b3870031544f4efd4a7768a703331d2d27a69be84f04fc0e807b45aff675176951fe2c46f8cf3dd67788499d8314037fa00dbd9f4d30346d849a185ffc518e2b43ab6803d343365f503476a329427f110e172ad5706f2f455d35878c9d1d9cf1aafcd1d8da97d518aa4dc47b3c7887a34ffbf31ee8d073172171c70ca53dfec4009b3fc269ef485f2a321f4101a8e634b37fad3343f2d8b4ae5b9bf890763a73d53df13bd19aaac6b24403fc5b4af0cc13bfb5356c6fa9f8546838372848a3a2d8f7f928baca632f55529d576ae5993dc6722b4b9438c023f13cf6e49dd8e4527edb8896bcfb220e2e8f72ae0cdcf37722a848fdd	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-02 15:14:20.532
h15x8kh2g79phfm42fvzsj34	user_admin_001	admin	create_pupil	pupil	0000	11411947c8349ce28805ebe2:385fbfff24326c5ac106cb73d7e72480:1381af0944900baa835c5388f3d6013f0f4151c4a6392ce1886e3b7d9f7a304527d8ad9fdf82712cb8a8146f85c0485bd359f9c07025a5bbe39aea24d4a9acc7a8e769d5029ae385c88b	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-03 15:24:17.391
l5mly0f2ovom14bbecfa82ha	user_admin_001	admin	add_pupils_to_class	class	d9j84yuyax8dyz050pzas0zx	994706417408ec89952ab953:aec97b20f033a2288746bdf18a924cc5:43991630bc796a9f28b1e66819c612ffe9134c933115b522001e6b174fb024527d5af90d9af48b0ee7f194913434088de98ee6f00f101fd8744040ae9ff0e57973caec17b080add2	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-03 15:25:13.528
ss5dq78pbq96i73v6am777hy	user_admin_001	admin	update_class	class	d9j84yuyax8dyz050pzas0zx	e916698bf645737a548d2174:487c0960b564e5b4523199c1d73d315c:1b242dd5460641fb9585dac8422d92fdcbb55364168477907132beaf3675e2dc1e29d2c430227beb6d679cb6d9e6436406d739cb062b98d26050545231e4cd999270ceee63106f72f522f80766f201f52e31525f87d1648836187bccd43fae89061520f96752d30fd94778b19d9f4139c680828352ff0a9cf1f8ea05e98f829b688df22884193d52a661964a3bd1e5ac1a898a4945ec6a3ac8d47ccb0b02e20c33274e1a342301442b0f84d33e8ee6c6fa2a2fdc576bae	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-03 15:25:17.423
xz7fm9p857vrw6uwjhx2z6it	user_admin_001	admin	update_comment_option	comment_option	bymqwmprlgflgl5ps40bjwdm	f847066ced522321277f9596:4953f106cba87f41607338e039c0434f:19e846509eaad3babef8088d76f0f980dcbdfcc08d6d1f426812b36608fc6202181bcd4f8282a4cd9a45e541dafca71752234c83d2995a96e3483606d5344d180dd6c9e406f7555a3e70e3b212eef767ea3070d6659890a2ccc9b35fee9c0ac25c41bb6710bef6ca2bb6fb0a112f3a851689a37baf768c32a0c3f9f093eea845f8a6451dfd0ede3466d1bb5ee8bffec9e760b3ff75ac10d05048f8211533491c6661dc47b2b5aa40f1a9f6815cff75a446ec89012c438b0b996254f8192b4c44c85b7e41f188edef6285aece3f085987a24d971a038a5532867d5a5841ef50fa06a9239015231a813353ceabc7448f2c97f263d328a1564e92491e2652c857373a9d0dd253500cfa44ef1270ba259cd3e14b650ea6fd44470ebabc48aca92323a4ee54171f0d2f066c8727ad7ebc8bcfa60fcdcc4cf4a8875cb13dda9bc13e8afb6aed57076b2aa89899d3cf2fb5c179ec2457d021df79888b6ee34902f18b1a8f8484788e64468b9c714bef0393a2e3363cfccaadfe48283262bd4f60f2b192a081f3f2dcc2a1c539a6cc19847af28de31c11e4f7a3ab1c4377fca43be0f7064fe833dba14df3e3718c31f93c8e8476ccc396339f25d9fc411e959b0a0f62ab483312f1a030fa68c5c5e7c147704bb430eb04f6242ed0fdaba73d5c809c96ddf1fee042b82530fefe55985f8a7114afad06731e954caa5f84a53b3562a18b02799cecfdf14ca3c75ed8b47ca338813343daf6d866737560a4df2414676a2a993280ef8d62cb332544b61314087a403d7308156223dc0453a399bf5a7f09e1da33c3b08c248b48cae242642bfd8bd1c47e35540d4e114c7ce5cf3ad18ed18455b411a5b4ab6f0792c220de875c48de2bbf91df55173168b7c745ae93a0a2cf99b891809fd970e4ef62fbe92032ba035b2f6d5128a4e22a69d3b0433c1939f0931fc163806484ba29aa1906937ee1c698149115cf787ef4bddde74f7a49101278a51538b7f9ed4d6ee00f9cbc9415209bb8c181f3da450d3a03443ee192bc38d0b170582a232255bbfc2c2f91808892674cadd0d0e4459433f6a56cb46a312efe3a896644bc867592311575efcf1acb2849a148828cbbda8d2bd72a71322154cba7539808c9058fc974cc2c9dfa8af1c7379265eeb3c22cee5f89261101facbfbd0f57015aa903112508cee2bd5b4d9d912847f5fd5c054f2cfd94d7e5a18b374761cf752e82b475762cbfa63e6e181acf956c06b673e79e3e497b0346ad9f9770f04bd46cca167884a6ba4b342e1bff97ca152456dbe18c7d435c9288bf6c15bef212d11947a2d004e222c65bc0ba7fa8311cf1a4eadfdb42f704fb9999271e49094745df5c367168e52378d048ebc8cee99fdbf3172241e7e9346e0b6ac39c305b1f8ca41b841f9a10bc36e8a814ce88a6e159f76ba4720e850344f04fdd3d9cb1a056a5c67cbc3e190bddf524dbd9ed0e61a8459824e20343203eb5387bb314cedb9ce58e119296cb4aa4db02e1a199ebebcc1b8486abc0b40c71677e98cabc13a3f66b9768b509acd6d915a2363a3dfcfcd6e3208b45295c8a35e	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-03 15:48:52.492
fdqltqrhkszgtf41g3lherel	user_admin_001	admin	update_comment_option	comment_option	ll88auca6utxs8je6ivdds4k	be3f02356b80571297acade8:38680f1933b7b1f1f69b059c4fff0930:dbc29f268701a98df637d8a9b12e2e16baea0f3c054915009456cf9e51c0c7ec6d22b26594c91e4d4884acd90857fa7cc7264a069dbd70de4b45d6281c57c76c003e3cfb331f92681dfa14dc2051328ffa7538d09fc401600b7a3fed2abbc54e229fa518d22e6ec9c849c51f97314b258cb836c215b96420dc3d95c26acdb672b3816a0b34ffc756f83ca0744454861699ac83f7da72d0a3cafdbec1af4e65c13856bb74e8184ea2a243295019742e5f9d3cfdea32e0ff6b0953ca6bd77392faaa4e7f68e9bd65958442a0095f7121b65b1226707ed4c134b4c9be728b32076c3eddb941848c62487f9f2a1687ba25b78864b550a98163bd31f4ebaf8b453755a273586fef8c24d946fd58a021cc4e0355391e549b3d9125145658b17239b9aa29bca0390cc524c0670ae96a174078e6508a2bbfcfd581be65875e250d6b121e712bf7f92847a92df9bb641329cd2e357f42c20a33600143049ec160ddae0e13040b6d4616b9b06c6c450989146a531934e896e711c03901428d6c7b184e7b534a507ffba2f3af8d6a4f5d4a89bca69eff84cf5889ad1142085e22817e56de571d2fdb96e6e8ab66c2971fe35d0af7d0ff17f7c5f3e95722e7482069730f2c8174454db8046e0c2e32719bb6911466cb56c9a92aea0bb13a132869fd3067f2db538938213d7dbeced41235f9cb5d7bbbcbd8e54cecbc2d72a69929f78aa5344aad3f454eef4f58e971ce0d74cc7032f14209c525fe435f801d9f020bb4dbcb70db32610b69e06bc70c7d7296ca20948f4a09ea92f704c7dcc29cec0adde3b1fb435fe4587cd762281296f905f99b9761931fd000d38590d164b78268a76c744304493e6b219e66f55c1268300abf250f3ab1161e9e30badb9174c50bacee0669e1a02b9c7296f159bd9c8f1aaaa537fca4207ca2c9b209eaf419b3cdbd0f9bd2ceb081bf50805151187b7c379d64388508ba21893098e203e2b5733907ba55110670c23ae564ed5c1400ce3ea5043802486afc0c482ae348fbe6a92dc56ff091ba96c4acb8cdcc204c7a96e952ffec288126322ca7b2da18eebca0c11196cb05690d9c5380324ddbfdf03981f7cf066e8c416a8f7e6ba8eb478516d7bb2d3a724ed59b405000e5343bf2a965c0b0e870f35efdf6e17fbb24c2badff78cda1d33954d5ffc187fe5b91371498ed10673bc14935127832846a0c9544b02613e247717ba720543231ab97ad19fb627356584292ee2820ebd464fb1d24558fdf937898a860a605d7dfcfde9293bd29bddfd6d40d5d5ac08f25c	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-03 15:49:40.889
m4xgrq7a7l84sg1glsdsk567	user_admin_001	admin	update_comment_option	comment_option	kzzm8ki6qgg2n4f50jadnmst	ef99b02d260a0f17a89318a8:597fbe9ab0eeaf3df4edb26eff5bb876:dacefbde3a7b31ea96437a7069f1eedba77b5bca7eafeb30413c457fcc526d7443a3e58a6c48a0787e21fda00938cf7612ab33bc4d6a5b7510e338fb5019eb38603b48c7910f3380b9cff90f065d4964cd97068e877ce1dc176ae7a1035ae88f29f963b124a4298d97201c938eea692acbf97ffe5ca37fe269b1d41bb92a501fceebee9174654a69570b9573c1f5818289fd0f76b5d64252f3f636ecb31504e565d8d23ea2b0cf462a886a5ff4e5ae2541ada8ebeecefae49c5ead305d8224a4cd41539b15fc349f9d67556e40db1ed80b60f32760f569374ee9b595b146aed5cb808196c1774fb3d9c5a75b276193f6b8bb1cd1ba5c1144d7b532035d315ae178d5a3596e0d97ff414daf4bc227880c5498966ccec7e1a145402d65123acd8520395ec4bfa0b6be56ecff704daa6ae7007de80333e7da2ef2484c184c513dca55dd561e7c1dbe2eb272bdf3e6513a0ed8ae694b0aee73eab91a4157ffd94f9d1ad1eff8c89d6090a41ecf6a18a0be81e53fedd8fd7654865d4489a486c46fff76d47fb9796733df898bbf4de45d958f2ed86fbf151957248d88eafe6560faf7eb5e15e41e421df885ffe01ccbd4c3a58940021f783eba2b915cb085729f7b32a1e1e5b36586098750f487b3ac32408dbe370caa5db90bc1f0ed05fac03ca74c8cc1dc68e5ff873f6ee9a7352567b73fa87c4d727a6ce7edd20b8917cbf3ffc6095618dc6cee253b0037ce2344755ddeb7ef45b2b860297b61eab9ea508558603d863ee46c47abf05ca1152ecb1dbd9b6e336da51a4adeb0613203db5c2f8f0b13ed67ffdef3c9b09c2e68badf394e58d827d54f73053fc34879ee997158349faaf24e4a24c7172fc7d870209ecba14e9522c3345e21badbe250b894d96fbfb26f56c2fee55add8d05143c73c1cff6e761e6c629fef46cc90ae289ea17ba6643134fbbe88d3b0443515e6eac506d4abd5b795d428854bc942296ddf463c0fce2b2a5c19fd34f1698973256c308b27d70f46d7f5b88e96fd2bb7a97b2eb962ec58fa7442d75d06a90cca8d8337364174499e88ddb1504693ef7e7a5614617ae57cba4611f5f4961e2b5f22a0658e3b5954f8b3f48415d19acf0588017d00f976398203e3842c0258cf0e79cc13a406d1581e8c5b9b3f3e0a50c88dae63949a50dcc8b15f5c4ead15a30e70ee5dc535aca81e4a3e374e10d7d072ea0e337f242d016d3bcd62c1c58331eac5585f501f80728bc919d4710f50f6c3c8c0644f19c7dba96e36132316a4ead3a417e4c4e7b13f95d2aea7da28bbd61baa5f129d3aa835aa2d2d8e0b08310d669776afb681f5067ef8c2e456fbbbf2274c82ae17c5824ad6fdd709e26dc6ebd9f597535781b0a47746dfdb4d8ec9fdc6f	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-03 15:50:00.606
c57t6jc212y5msluowvmcjvc	user_admin_001	admin	update_common_comment_option	common_comment_option	cs3r0ad7jwxtv9l4dfvvukce	b738fa5fc72212175c99b3bb:1bbbabac7d06432b3d175bfd34d0d74d:9d071bd06136613786d4a21727ea3b563b19311096f7f172d2bb2d933a1e87a0ba9e2f09f27efab67437e1c673c12f197bfd3ca72f00ddf973e403ec15ccd1857ccbd14e297daebf49b2f3eb65439f5ad35387b0754b3fac5cd19ae960a2063b7706ce154b9b17be7d903ae1bf26013aa13ec26e85fb7ae4b2895b597d60381b6dccc22cc139c6375f4d9892a50a3b408f365f48b3361254747809fede535c25500962fa2611dd7b67ea4d49ca6d68e150a88c07f7ba4c455bd25068a98c6d9d2069ba733ed55f476e4f7f672012308a62e684ffce43bb60375017efc7dbee0e6840f803a10b12c368ec996db30e1f69e0d330659f6bcaa41dce1463e2	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-03 15:51:10.54
rxs01vuald3n3j2xu70nnq7a	user_admin_001	admin	update_common_comment_option	common_comment_option	bqp96t6yfdbs2lzx6cdnf2u7	65bb2fcc524cd2882c65a495:b152215a13afc0873f0e5010160f5237:91b31a64cbd8656c19782b5ca54310412973b7fe8405d451ff135ea5049e65548387d43e9063a4ef301e48f5e2630b8c8969a5befc838669bd0fd12da3928e5d3181c47d2607f6fcc4db1422e73e63460969b14767c8a09f9feb3f2d89bb5538118b71a49b63245f21a14baf551139f0953fc0b20a36b4621c9d0501dc125e9a9e6badd192cef0f13dc957f54a784748feb5e472b0004cf4c7460c7fb07871077ecf051299b8b299b24b951b8a1adbfd662383d714b48635ce876d0fc9820e535512df7452e356ccf171db09366b45cc8498c3d4628fcea56552d03110ad26ffcac270fed0cd75e461005f28c1f2c8caf41a732fe895ac7092d9632097b036a30810fb71557416b749ff06050255b346b46eda6795429dc828ead77bb1b54d216916df2a2122325964cae5e34f57e4c33798f7c40f24f012c03bf4e8f647646ec5e27fe7218faed1c1fbaa798cbb23cc0e53953dc23731502fb15b9a73347ceed347d7ae0a80d32ee1f234fdf1fa22e84e31cb2aed58bdbe46e23db2ed755be8eb44eafb8ca7c68d0c28edede3f18dc2029440b07fb99b4bdf43e1b25135ebcf86bab5c2db19da0856817c8eded6860620bb420a60a425e309dc49c8308fdba0daefa0a2fee64ae4bc1e9193eced8ef3272daafd8b17a4ac1ccb7afec28c644dfa270e5c83a9494ed9d3c7ea0ba4b49cfe4cc2e829799452a87cfca72689aa3cbbd34bb7918b91d7615e3c073d9db286285a961f992ac6c3088ef2e217c3a096a1fc2e291e692c44e0536bb0db733dbdf32aecd540c33ac88b57440f699fd1f3540005641ab731effbbd501607d601dd6eab00c524242167d5b6844e86240bd954c1361feecc31d61ad96870f50a7a68c06c8d2748cbc29f7f25c28fd5a8f86bbc768c54dac8fd3c3aeaa3749a45e56029280fc56012cd92a793818ed0932e82d31b6be49665a63fcdbd4e76fb0185353a27b16b1064dee602d7e13bcc6fbf1748c7f52b7eb1bdba37519af49cb73d0b9b3c9d9bb2ab59fbc31ccc820be4e867023edd36c51ea501f864453fbf3837bf9e6cbf847bd058fc495de8bb21d317738093c93254d495b5c7ceded70ca3d1de446c6ba6f887adf70180fa4cf33a51af67cf5241524d5062694076df7bd585d18206945de85a4bc5fd366e85dc839fe33a5c12cb70e533b2a117ca2573519f2b9577af43ff51f5877e20f0ca17ce464b1d0fbd0565ba2d704b18bd992623fdd7358420975418633e692841a0163135717010afae00f95cc9ae3603254b3888f2ea254e0270d5a768c09da51814fbd5bdcf290d2b1a40edd7a0840ac40a	51.235.163.8	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-03 15:52:12.33
jvgpa8yd19w0u9qrow21wn2q	user_admin_001	admin	update_common_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	560b555aeafd3137870c966c:12ad3e2511ec7147e2f364a8ddf5656c:40b678f71e4ece81585ee1ab6e07d8fab94d1dbf7ed566c261c2df95a38529cc499951939e3d9df79d718b0a60b46046bfc9b72706be8534d5e6a8ccfbbd037c5eec7f7b31a5ae1ac2532eb0d761004d7cf11112beb6cc78c9e79c811f01c53f1bbdca90288975af82b904163aeb4fe52a198d522077d10397e1a90df8859f29c37a7bbc27a6	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 07:56:19.538
xe1yrrzuelg9y4onfbd8bcw1	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	8070bbee6b9b968ba46a4e11:a28960d61a08f27b9aa071eae00d2225:592632d1c59a898cb5bf714ad5288e6d9ca1129a222248e109b4940009e2c11de70ff3b04dc061f5741d7673dd26b57bdb81d9b2fa3dc4aa045f712babaa2191f927e8ff2e4147665e78f5d5a00cb06506d6dbbc73f476c893274337d12756fdb66519d7f0b676a39927f3bf43b1e4fa4ae7cabf2ef9e0	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 07:56:20.771
fj2znzqyixbxxfo2hoaavy8q	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	750ce9669a483c4073fdb65f:3a518817ca470af4d8ba64c98540ad80:c18722e538c815268ef18fb9ae2b6eceabc8f3a52949f32401676fb5392dce1742437708bf38f9e785630da8d26bd6eecf847ca5f7bf5e5cce3cc9314d85f267e2f9cc00ae8a3b783a141b4def100d43b7190b1486f8cfa310d55235b5e5ce141e09ecb8b6da138da6ee31fd7d1c87abf4c8c0069613a81a0ede	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 07:56:22.969
ccuk25wsfczxmc5l6111tx32	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	50141aaaa31a2fe7d1bc18e1:47d45099147c878698245473b74ef4f0:c2a323568aab16b504a70fd607e780a774c03f05aa3c4c1e658ec770f232669f662af3f8a7d2171afc39e4495781bba667a53a6f41d69ef6aca83986f5519a91570b0b3a964810a4bb90f1cef30385e99564e37c5bcd55761292fa252d7cc8af1e638d76fc8b370cb5eded25cae7058c5dd49bcf06ee602f0753	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 07:56:23.631
mk15c3sbuivn9ghbl9ud5f3c	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	a078cb6f549666bc2ed5b942:bc835c4b1c1ee75b673f71cdf166f7c9:5f73ad955b9bcaaddcdb1ad46590b14353a30ae9b48d6a707e63c14b77bd455c0b8430e171576484587a8a20dce2c2eaa1c534bfda4489d7294aed9b9f9bc9f994f29f8c2a723f82cf92687506ed6989b5ff14a7f23498d6d87e1d7e60ffe16ebf6b914b585192074f4de6aed3f5f9040a38892763ec460184b7	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 07:56:24.575
v33d24g8t028qu3csa9v4dzz	user_admin_001	admin	update_assignment_comment	assignment	d9j84yuyax8dyz050pzas0zx-0000	ee01ed1ca6d660981e9caf22:42ef13108f97fc9ad3f63e38210e1ea2:0d1e0561f14b8065f218f10ade1bbc5eff12909e9078042f0c13f30e3a247da443c60c98ef9d2f1e5bfa0e9f418a06c3bc565a42bfe41a290a1274c46397da1c20b430309cc0610f6f949b4ac5a44a603c216c2d39fb1cca73642ae50bca54cfde83ea4dbfcaad5b5a2754b9209229e638b9a2c574f687af64b09c0032dc41e1b5b2dd5d71e31c7a3c	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 07:56:51.658
p8zhj6j07hnp2cr6rc7vo27i	user_admin_001	admin	revert_assignment_comment	assignment	d9j84yuyax8dyz050pzas0zx-0000	7afc502ee68f9f81934d3373:d73ef391765f0783b86d021a36e58556:5726b1438e890d17d9e65a04e589252c3a925d252f1abc65564b8003ff0ed95e346200898d44353a86d88aa99f	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 07:57:24.604
ria94wal41k3ihv09k4fpvr2	user_admin_001	admin	update_common_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	2ca4bf8f6455d27056f5e1cf:be36bd74d7f20213cdf3db59e4c65142:04155614f97da65d9d9faa0af6137222504a19dcf19e5299266aa58befad817863a146e0f2aaec288c1022d7608f69af8d5d56344fd6f649ce28d09ef01bd6f41e55356fb6c1b898b79390cd66c7c6a5c19ae6fbb9e13524bfe32882b4baf63bc9a36a44e085c49fd31511ef034271663f0ef0afa788c4f8776270970b5b8e94d59bac2366b0d2cba965	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 08:00:01.07
luns5emuojyd99b9qarwrljv	user_admin_001	admin	update_assignment_comment	assignment	d9j84yuyax8dyz050pzas0zx-0000	f12061db9d0306d6428c25ab:8adab3bb5afc9605bb93c1ad974da901:740cd7c02f066ff1c994a7ca0c715796ad0b5dedb62cbf5d8e9adaceec0358eee56d26dab24c183f9c9da22922b4051c27ebae1cc7cb2692e4f51905eb9bfe7574b734ac18f2094992d969153d8ec74e3338299997dbc82bc9631c1effec45224dce28e7db2a7e1a185d38ffa2f918a03031701c11a48444dc79c156ea8cfb1bb2b5efbfeac42387c2	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 08:00:19.812
u1ynw92i6pw5q2a7fsr54eh9	user_admin_001	admin	revert_assignment_comment	assignment	d9j84yuyax8dyz050pzas0zx-0000	3c02e84565c6f9805af8506a:690ac09d0075839b86d3b082372ed438:0f5093b109400339e235c096142f2af2620f9dee623189ab10b10db27b920d40e74d1011961f8318d4c1915899	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 08:01:21.638
p4tfngealuvi1g7x1xu3g995	user_admin_001	admin	update_comment_format_template	app_setting	comment_format_template	5beb55087885fb784eb4b6e8:e0f5db83a8ac89fc6fc54e022111422e:f9c042f971eee5efab566b5cff91a3d537295b50276bdbfea76551131756e327f7511df35ab8e6dc700c60c460a79cc8aac6b547ba7dac973eab24d6cfdcfe506345b0a0e1be8a17d59b841cc6879db863d25aee8a0c741dbcc1209b9c23731eafed1c84230bc3f236545f0f85	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 08:03:50.79
q8xhqxnu2d0gtxmr3i5kmovb	user_admin_001	admin	update_common_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	864858f55c6902c5e5116b16:50e26616239a291aab2ff47a8ba142cb:085600a12d58bc5a9e42876687747d93767f32ec45236cf608163e152279fa72a2b8a75d915932bd522a945e475644ca1dd338f3ffb62a43833def782277f3ba3d35fe344ad4fede9c7594ba83bb4352b04d8c5842f1b6573136b84f6c44c2094cbe5f05134e235a96f2172d2dd3817b5a91366c93e5a38ae6fc6d93d1d122026edf	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 08:04:45.129
mxl62mxxcw5s4flc8knxlz9k	user_admin_001	admin	update_common_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	81027d57acc238bb7224d05d:4b8ac3767e7bc2d301249ba159310143:3e5a0ece4f3da776b972de3574ca9bf29cecd5a4a00f75c7fa7df91b8a1522c54d0903dfe8feffc356c763166d52a0b2dcae0f5c9388f32f32fce1c62e49223f7fc295b566ca896e41333d9f481d0b34f7dae0d1e65d016424a6cc2e66d807914bf7d783e604deb48f81a218f79d7b168b57aeac596d374092db5ec59798d80275694ea1	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 08:04:45.655
p4ojo7kja1jh87jq73e51v6q	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	bbb36f105f93628e5a3c994c:d476264c66e11bef7fd62ee78ffaee69:887bcc15b0778193792bf7973b2a6c6c13f1f7c929cd5c15954bdb850857c9984228745a96cdde1352ac5bf9a121596ea5b46def1be6cb2aec4a4c71b1672a5c55d48a2cb8580bf4a6d153f4c34f6834b114a52a90a3339316e134ed5e0df54c57713abe2d288a3793f059a9200bb7f9bf000219d8	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 08:04:47.535
k4st6yxda27fk6vjwupswjnl	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	7347f35281480ded5f73781d:0712bd5a7f5782c74feff3cd0dc22015:7ef85a2740fd663a9c81e96da27d1b71cc75ff74d0c5e118a2c0a7e6b4885d546802493340677099668fcbd60d3672bee76bdb1b8396ae6a2f2d8bfd20a9fd361539aa966033a4efbe10ffbe23cb24449565a847c6a9cc76fd5c91047b04ffccf9d820ce56ee1dcb570fa985291dc6b20b8a3dda554dc7	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 08:04:48.493
xaxdnkdrwbgd9evx5d7und4f	user_admin_001	admin	update_assignment_comment	assignment	d9j84yuyax8dyz050pzas0zx-0000	83acc77760f4e6f90c5b2f37:6d2437ffcd6daf2ba9ba2b1aaa88cdd9:32ec41fc923ef4187af8f97bda63c121bd13d41ecd30943f446f1a59d64663c859eaece0d45049faea17ac70d617de71e1809c4ecb77276c197657e5908109d2d27cf63b4801f72cffb468be7bfb756795145319dea5cdc7c37df97a0b8a1732950e92e4e8e2a98b2ec4bba5f05d65a6e3d2d338a40de753c08b014143ac3806ca124155ad071d5a2e	128.234.122.239	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 12:46:34.768
fbftvrh165oo547w2r470tu8	user_admin_001	admin	revert_assignment_comment	assignment	d9j84yuyax8dyz050pzas0zx-0000	a92a18963d3f1206dfa42a04:daaf162233b4cfce9824ea586a650c13:cb83a9702d999d4c4c9dd5ca4830291d9f74df8854958cb8e5ba154ec091fc5c2231f20fb2cf23879a1f6998df	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 14:30:59.28
mo64gecxjyh3x6g0hih70adm	user_admin_001	admin	revert_assignment_comment	assignment	d9j84yuyax8dyz050pzas0zx-0000	08d476f515d039f78c77f182:5a9df6b3c0ad1e9facdf028bd8f49c3a:783a95d0a541530bf9df6196a17fdedd1126525fcdff45e85d0e3310262cb745faf450eed7a0e372eb431961ce	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 14:30:59.314
o75rk17m6vuayirjeg39wy3u	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	2f2bfc06d3c7d2a903236cdc:21c3e008140be91cc3ecb960cd6a27ee:c580b2afaf20a4a4f0e7d125460380f24bfe8a2c17716e8568692b7883b35c2e7adfedcb4f627f3dc6eb4c9aa5a10a0aed1e997c9d42778f05d08578644f435baae63e308bfe2da832fceee985ebf564ec75d51ce92807381c7f48e17ef5c1e87baeb7ae1ac45199487faca4f4453a08902f9ab8e0f87d4fb4d288	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 15:34:28.772
byhl4seg5sl1qqo9nz6jqgfh	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	b35e5ec8b586260e85ecb4e7:3796f1a497f00817cbbc2639d0a68056:90051aed86378053c70e37e63e33ddcb847a3845400bfa0cfecc08492f7a4328ba9e42728dda46c66e0e95a95ccfd1077f44ef9b5edb2582629da3daeaebcaed79a9c9cd6ced8409f2b4d89a36db8165d8996eab7572fb3be5de6bd0e1dda276b5fb97442c24b76ba435b84717a875342830851c72cf96a56cd6cd	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 15:34:30.227
dg0uvoet5oer30s13855wn4r	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	bdd85b2b2124ab86b4cf92f7:59bdc61e153fb68f820b63c47b9b5bb2:464160a84833b059863744d0737df2f4c080ddb94d30b76c41e248452bb16d749eb4ea0aea1da0972f25550945e20eab9c7393080d301b8fa3c7505554fdeb96052bb9959c2c113b0a0b6a0772a7b004e595c78f81876c3fd0831c0209e8c000e791f47d0f57247fa309cfd5a2c793b9f5d3cc8ac63d59cb99ebff	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 15:34:31.481
mq7derzo4cvzaoisbl87dwow	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	5d9a5e3b42d5c06f9f46c076:722657a8ccfb179918c4f48897fe4a65:6aaf79174d1305b33eeb3b2bf4b135caf7f43629763eabdf42e395a3a8610c5bffeab30aeffd0820f33c95f03a59be93147eca6e6e2f346504efd949e671f88022eb5b2c35e160040a8582a12698b5388558b19953773b247e02fd617bdf25f68ce5bdbe87a700a6e12de746256faf9ee8f4258ca9a7c138485df2	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 15:34:31.973
sdi04t1ygy1ubc8pq8sqkp4n	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	f97d68e82d0a43ffb0848d74:e8ea70ca07d97a05798b3b50d5cd20f3:e678b5cc39a55b3741343d9d62b88fe4b3a23b16578122ff95a664c5bdb63a820e1fc53a02059f341c623eb93963938f7b3a5d71abf63e4581440ce1d38a16cfcba8ca86b4d10af1c3bca563c28494285234f634cf67e0d810965844eafeaeba2201395eeba7cb047d5af372b025348bf203dd66f280892f68d936	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 15:34:33.226
l5au3e4lmbzexjk5xg95h4zo	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	acc9c3065021339acd1c2dda:6ab169c8259e6ede7591ed71ef3a1009:81c4c81c9b6b0fa50d005ca25fca9d29f55b55163061e2c840077323a1299e2ab4256f70c3fe89bf6425a692bc88abc801378eae9853436e4c6dabbf11d0e47f0d90c5780c708d0a8e5851fe7d892b51f5d46272873bbba5198277da5e3708aa1e42a8b1e19f42311455a37c157f57d71e1a5cccc96ed0c0bd7f2f	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 15:34:33.705
hgyd07d8kcoe56slf6hgyzct	user_admin_001	admin	update_common_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	f98ced6c9d8a9e8aa8e25300:bbe1addf168ba97acfb612acd3ea2bcb:340ca7abf803aba49dff067c9e3df038a7e941776b7aec7b179044f207a0e1571e4ac1721b0f4b66002b0fb19c6b156dafd18ef601b07f02429f944e8925e3818439e2eab2fe3fa03b629155a45c01112239b41d3223e8c821e5001c34cdad1938b623803d829489b0d5e565960f63a161b96282a811c2b2fdd8cfcfb9798ad0b2b61e14938d97	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-09 10:50:56.082
r92rpebz2843vmio1yi64alr	user_admin_001	admin	update_common_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	c78ad02d4b352cbe5943eb16:224810298bb59eef2ac2dd54b6c94f5d:c0261a655df90043fa5c938ca39b66d698877dc92f06debe28aa55fdee7bcfa6e08a3977ab138ef4151cd702b6e79fb2ac2a4cae7743a05e357c586b81b7cd39f930ec2696cd43cb5b6c6fdf5090058a390a867352c5e95ffc6821323440afe13aa35f001f5b45beba68623a8b035e2d7842a65aa8e07716d29b722492144a46a493a05f350ed2	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-09 10:50:56.665
d2s6ee12iybto4d9e19o7nxm	user_admin_001	admin	create_class	class	e839woksk6uecerw1beh11gf	f23491a4b2955a85631f5cab:95e709450df8f4b84c9bbddf8e2efc7f:39cbc73f10eb7ee977bcf7548b5a33c4bcfd97906206875d5eb19b7a1c8a9b5cb6876556592fe0a294792250f54b6e1f6188b28fe8d0a59152cd44f310ce72cfedca508f9f1b76bdfcb88914e177a11def5bb531b79b8f741cef26599bb71824cc270c	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-09 15:16:44.737
zurga0he2xbbewb7rolrqo68	user_admin_001	admin	update_user_roles	user	7127cdaa6d5aac38b1ab	1b9704b26115fe365c50951a:57b3d27d10ff5c7d8da4c4710bcbacfd:061c74536e63a60931e07b506728526817499ca48faa301c750f6486e0bf9e01bfa9bc86be3c1b4c7a389c8a064e265042b1fdeb3570416e5ca3404ac99de3667cff37e884edfa9241f8cc325abcfa70189eac7b97ad959e18	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-10 13:57:07.687
l0fh80w37non6b1pkbe5cllt	user_admin_001	admin	update_assignment_comment	assignment	d9j84yuyax8dyz050pzas0zx-0000	1577b2982eea58ee60b1ce50:b1acbf66ab3cc239f6a098418fa6cb33:99163e73753500a17f76464b28c056fd2d228254e0a2e966d69f137571c65a4fd68003468f8bf0df60bf27341ddac6622e932eae4b08dadf75a312d9c35ab73bb35d00dc55cb0e9b31b34226edeccf29d2616c004e7c9d2a19a38129aa907a8e6c2c20b1b7c1bf61f6e87930eb54dd35e959d64773d10e092165b9f7951ee07f5e0c43ffe5d2f4d68a	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 14:36:43.342
pvt5yicpa4vz5f5mmy03nkx9	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	346814b56ed6e7862cf2edc9:95c2e21effe7597cccd3a1fc2cdfe71d:847ca71eba8cbd95a9e416aedce2d66534ea4272928ec1849a62bb3af64ff54066b1d6618c92a48e1ae4369fbd1175b3afc16e73077970e95a9d2d07d790ff5f38b755fcc5f6c336af66ee8b385ac4b7e608d9f1a415a3e1834c9ab72e5b5bcc6524bb8a450fae4d9411da7f7b8be8c4acf901f6977380500fd80d	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 15:35:46.321
m2gd8q6ace6coi4gygl4hd26	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	39b75d28fd39d991c45f8f34:3607188780a85e00ca282ccf2a236373:de2c7f5901cd21787f2168a29e67510150e00a0a87eea2d55699c1a261727b2c3a26b27af50ba4e48a3f7aa1c02711dd96ee4a1bd4474037cc43ba4ab0aa334f016cbd25714b36e86b3d3a761050e5418dc7481c3664aeea067fdb54b02c2d8df618f25ece05d01160c9d7b09c21c2749d83c2a61c66629bf1d9b6	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 15:35:46.964
o251k43j39xx2nndp4cksdl0	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	369292034ba821ff44e85d77:d67793757cf5135070a79496144a2c92:fa3eb2da03ce6f5d68be9b8867ef010d81a0cc26abd5c946bbbac1d1e761996d559a191692a06834866b9b18a9cec267f507e7f36efef87ba1049e68e60f064b38d08ad19992eb23acb2a58640672a5f5c13e7e8a05dbbdf31c374c328d2d000a9d04ab5d7cc4b6387e4dee912de7a315e7f0a0f569818b8c18bf9	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 15:35:48.192
cixc9zj1x4evflqk5rrvalqt	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	7dbb4600d77cc2ee68daeebf:62e72a9eafecf28bb44e7495c0f0933c:264ccc3de5aad3aa43edb06e6496f758f4299afb4fbf9cca133bb3715e7afc67452bc92cfc89654fb3dbd71eb560b2a188551ae7b3a9de2b4782170672d8f37b8da179034cf4d0dfd957d353a54ab85b442eb6f41617ae4131e8b1c62b01ab9f238eb8fb87070de4dc0c700b27e33c5605aa5dbd90b056bdea0b7c	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 15:35:48.876
qrfo07flepx7zd4hh3tuwff9	user_admin_001	admin	delete_pupil	pupil	0000	02ec9445496f837e7cab6a69:4fc684c72044a9d647a599fe56754ae9:9c433df30521912adb5350553d74ae0b055dea3be874bc902c4d842c498dbe18cd475859386471cc51333fa05b3151b3e51f	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-09 14:45:15.598
sx8vodxfflf9y1ae821g37k5	user_admin_001	admin	revert_assignment_comment	assignment	d9j84yuyax8dyz050pzas0zx-0000	cb44e859c700eaeb6cc0364a:fefd02d17d2b89835d16f0dbe6df16c5:06814590594d59f6152aecc75f62ae6bd2ac55cb7c63506c5c3c39be829f478b3908ee1f2553910b35d52b20eb	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 14:56:18.862
j1jmzwlw7oybschxlyoyk7ik	user_admin_001	admin	update_comment_option	comment_option	kasbyby331o69w6kusa1u88f	976b5852ec785445a73606e0:7ea2b9f454e418d83689f62d28ac4a77:74843669058f057d23093061d687109729b6a91d98c745290c4ecdb0f8da363ebe4269ad595a2d3a68e51c967fed4f8996d5620acc8b5ce4157d1a26c5db9ea5dc9641f05b3be2de6dee87b08a7ea9baa65580b0d0e29449fb3ad491c19e6ebf1d7187e1d5ad200107ebb88c91cd2f3cea8a43ad9d1e1a63570b2dbb38d0979a750838b569756e48922073b681792cc2d27e70922889d3aa030482cb2ecf549da89c132732f6ed70bc286b994e2f97d423b4dbeeb5bf3900f645c1acb068bcd5ba6b82bf115fb941584fc1da1502720c267032e9f9b7249d4b13c5655d1d8f24cbb6b679b27be7589017f82ec21cfce3775f40f3e726960f06fe373d06cd2a55b9a004ea739a6c634a00f063a0dfe07b7673026727fd91727f4d4c663b26b68aa0db27c82697aba6d625303988bc613a8e5a7fb8d15eb0bbedbae905bd3b1228a198744eb709d525716303c9d6141cd264a9a81b3e0752da94f4470a46a6d2d85ddb6719660a557ee9d348ccfa6aa8b315e74a5818400cb5c9d22265dac8f7bbb8818408b84e16d679e947719b2978fdc64937e4073ac86b6dd39efc8ee912cdadbb18d178946bba5723922609fa805d38225ece90b5517dd221825cec367d0f839a4fbb6eaacf236c145194089a82968a	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 16:36:55.574
uw87tfvavqlpgiauitpyi1ap	user_admin_001	admin	upload_pupils	pupil	\N	7cffa85eb952d198965a654e:2387fb8396f33c02f5f1f3fcf708e636:a2a2f934ce5e71235505ec669fa7860dd638b23e46978428ca2337207eeb66300834b6b8cedcf7514892c4edd7516d1759e38ad71f5639	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-09 15:13:32.821
kh4q5srltlvjx4ygxhlat6os	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	6ce34e7fbd9649bde291d08b:55aa554139dc6924cf68cbb7ec13124d:9b745641f4de08cdae18b6ab523928016f256caf35b7b36d74a7f65f763047ad080a86df5b53d714ad25d46d2d7cb9d265d0045f9b300db6f8b12594308d552952ecfe487ae7b8a50e9cf8dda2c0d9ec6b95a2b37a180347c039b80ea7044008f178d8ab01e5a5416782b388e52374369244b90c81456bbd7cea55	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 15:23:22.76
ze0q9gs8dh4x8z1jt19k44zs	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	088d0734f9922cc93a73ae3a:334d11fd70443a4b440f20310d8b5530:5d8091e28816a1e76ada917c3d9dafb2cc047dc152f67e1d10753d491684c4d87672c773c80231d049af5f82755fce29ebbe4113fd8bc721e4a7b8563e14b9b066e76f9836ba8ac78f2804eb29e3793c4294b78018f25b9b193637558d3876b45110b8615c2393d72975a09e047c6b9decc639a06b323d90de7cea11a893	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 15:23:24.411
zvi3ge5d720am5cpzoyvli69	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	d2514854b9209bdf6541ef65:6dc446027caaff62932f846cfb2e9ffa:6c51e9117d153d1b3efaba8be6dabc93bb49685fe17db2dc2a55109b16d1e8b1f69249c6a8d3f370fcecc1064fb9209fc5007556615d4df675ecc719e9b02ae33e49c5431b08541d20764ea19838ad23f4c19149b8319c3d709938c598386be1186a20a9d15f7cda828690228ef4a67d032c43088d62a8f619dec61bdac8	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 15:23:25.827
fqs11ty32ws9to0915do62u9	user_admin_001	admin	update_comment_option	comment_option	kasbyby331o69w6kusa1u88f	8722899c211a9fe3ae5fdefd:6f9390ec20734196dc7bca4ef46325be:57034f8a1d0fda4e040b441c8b46c2b11eaa5cf864fa1a29221d7708e74f67bd4a355bea4d0343013b88d28ac14cdff4fee5edb3f57d5a7a780fdbe18b27002d290f958e02e4b9e369e4583310cd4e7ea05fd8d14ee6cd539bc6283e18fcc9d77d08058059708470b29fe81afabcdeed1de61dbc7040114a28f17241951d6bdddcbce3e7a7697ab32c1b73d61d413e7ebe8c04ffa0e0062f66ec66c215d6f084b359681bf9811bb1a4b035a9dab06ec949004f9fd22c08716ac71c996cb8b7a1950599f32bbdd218b8a0df4e326d6e23efdd31e53b58c1d5c4609a402c8248475ec47aea707cbff1fdf5edd29ae0b0f3bb3c89bec21f6343300e97c22f83ed9e7c4eb6091fef5a358072358b5e018cad4ff493ea71c121aac54209ff38c6435260e15334c59b29c2a9ab889facbdcd1b8577e17761e6a71e0ac151b68371cc6fd984f4250112b942138559ef591d9237d605f056da8391fce04ef888c3df6f2bc9709d8bffa55036bbc69df2dddeb81c3b28c00c03131b201c07c728568c46e003a01857fa54368b2ec41823913fcad28d08fc53594d724d573a3d1a159953c8e8cdba61ae07bc86b5373f0ce716e8a6701dd0e1efa8b9142a21de8fe4a1c0599613b2b6b49fff8f56d85bbc3fd369d22e0ea41b	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 17:32:42.919
d0e3enosu3s3jvrl8lvdukyp	user_admin_001	admin	delete_class	class	d9j84yuyax8dyz050pzas0zx	78fb82847ce6c492c4e0a80f:1006a32988ec1d1d30bd3d98262d97dc:ce89c57cf140300844dd344e5ad8b7e18ae8451c960588c4aa3819fc5cafbd8a01f0d5d5cab190145170e08347bbbd3bfc7d1a118b66574ac1575ef43d110e3d95	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-09 15:15:37.084
ht2kk030reixb1yuxo7ykinw	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	e17b3358dfab580e67054f42:2f8803ea3c5c091e6eaca853d2a20b56:23301ac4fccca7cf15094c34f17ce52f5051b9cce0902e562cb6b6cc20849d52b6c13f8f63857c1598e96114448521b04ffe4ef1b821ff8ea7a059d1295ff363104d0b692163cda563ad4d1880d3c3bc47b9971ebd5f5d78be00a092202d2344216ccddfb4ef7abb8dc7eabc03af2cc35083b5efa34ed47f7828c0	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 15:23:48.454
pz610st5uo2m6p58ep0dxs09	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	ac663430941d653b5395b24b:9f60515300914483b21334076c857925:8e4ca3e02788786ccbf91fb4cdebbea9a9a323589c2bb96abfd1985c233ec781e8d4aced2eeedb66b1cd10288efbe19bc783e1855902cd1d3a1a58d46a4eae9acdcaa2aac5ae191b951d98ccf1f72854d2ca62c8f9465a50a1b0d6cfe7d12183cdb45207019127ed65a92e9ed3d091832d4e183fcf613b44ccde8c6c3b1f	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 15:23:49.434
trvtxvf0r5nf3nc4drxes800	user_admin_001	admin	update_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	09e4b0fe87349a6fb4b8ea0c:8a8dd5c77538a7f829b6f5216c124e20:7a14dcc7539ec40f39c3cd7239231970605911099723e35dc5c6bbbed832f5543646e822f29bd66d9aac881b5a55a04be2fcbde0e942ce4a2f90c7184589eedd915caea866bec231023a71e2d40990c4a7532eea6368d57615b26fb7aaf0763b190de26cffbe5d22c5e2b885b20be01e57428784369ca9b2287cc4fe7acd	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 15:23:50.18
o28zvrn7j4bbw0ehlm1zfugv	user_admin_001	admin	update_common_assignment_code	assignment	d9j84yuyax8dyz050pzas0zx-0000	f43b6f212f6dd9120902b329:7d4b401149444eb02aa1beae309f08ea:55039471777d72f7d3cac29247ee72979ce2b8e5dab9f08dd2409e7c7df6bcc50320b775dbbf2d0da883dbab728bf4740b98dc41ca4f6951c030486b58cf4fbd35f8fb7bcd9060d79e9631dec643fba33cf23feb12651001fa7d7f0fa4e0dfab1bd1fb976eeb92f24a6b68943a7620e87a8955005586f6c508ad25c2d0f6f9db9a1c6a23223826d2f12d	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 15:23:52.034
z5dhtzkzxc36pfekdd0sbzds	user_admin_001	admin	update_common_comment_group	common_comment_group	kvyc23qpstalii62z2ytnrl5	594f6192673fd5dc86fec8f2:1792408ce2e11fd227e3a4f027093bc1:adb313f50fc51436a27453cde0efbbf416fe96e7d85b60d6c61d76b91b3477721a7340b79997158315ec6220ccdabc1e17bc69afa0f44b27997bd16c9744ca70e10423b219071c3d117b7e88631dd867112f51b062972287e29cf11e3e9479a53ff88168afb6ce75f501971f12520075eab47d0e2015d09eef468df0f15428f058a2ef6f57a788bc0a6f42c46700d581af3da4657a18f84ed388fdf01a749224d2384fb551e74199436a81423af353e146c595854ee0216d6bbbca44	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-09 10:40:06.473
xhvwwsha471v805z9pygai93	user_admin_001	admin	create_class	class	tevo6cuaymo8w54pdu6x64kn	3d50f3ea371fb44204ad67b3:e52ea8bf74fcc8436ca6b419d128aba2:2b2c94860f8733d1b0759071b232c2b23776616bd54a2400ae0db5da63999bdd90fd1d758302043b5396ba99ed3ae94fec077650c0b092b7960f89d2055d5406f48f543f2d6ffbe6e63821a640fd0a514d9290f95d96f5470725d7df0acee6bf480891	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-09 15:15:52.911
bte1h0snvz50cc24plhpi3jy	user_admin_001	admin	update_subject_comment_format	subject	DT	c877678250911ae2bbd62b3e:1dedcf347fa93a8f2e2bd57822563ed1:e138ad1867340d45f837c565099c0cbeb64e674603c566d1eef3a01258c095946306faaf2022d418f08ff6c80d415cd48c7b55b377f954a8289c83bdba7f4b361030c9533a9c04fc92e4a26ae8af289870b8ff141c213b23c752b2b0db4d5d53c349f2b40b6fda0b50d0a65ded21db39cc3ac0c82a25df7795b693a485b94cf1dc84fe3ee863190b787e101d575894961c5e3d06a4a647c176693cef26fa0827e9bc8d025f3e07a179dd45d080bc487202ce52c0e641	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 15:24:50.622
zj8mfwy9sx3td04ibwel3n36	user_admin_001	admin	update_comment_group	comment_group	xdtpcx4ay2u3o0mm9fuuso1l	8ad6812559036969a937c796:970621705dee9a373b9786922cb5e95a:4d43c0883ab754f1c3e651752ac0497f9521da0ba655d74e885fcffb7107b10d4f9cc1ddb0c6cbce0d41aed9b3c3016076ff60cc2b1ff6b5281da0aca33847008e6dc485b4e9b987841a3ea752fa2ba4bc135d3c681817c648499cb2fc7ba0a7c0929abdf64c91385175e89f8c4c3f09171e1b90843785ea9f52c64f9784a06f023ef6f38226cf10fc941bad1f8493c2f0438b91e717490caab0aa664bd2fea9ee2e8082988e8be9d08a696ba8aaf10c3c917144373b9f6c081bde	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-09 10:40:31.486
r1bnbgzy3cjbl0sl9aqgvdmr	user_admin_001	admin	create_class	class	yyql94wni13fb31ffe8pob4q	b34d1a3623a47ef947d23ada:5a410f87a6d4f68ca41463f9f0a2aa04:de7ff6b773026a790a44fd3c6eb18d574ee87910d6c45548b96ba8beefb5c01eb1f39e3bb9042c3820aef71a845c04aa37a3b79dcd0e972af11e043677f615b494140108335bff39de68249a2ea974d44c907e77e7b79d1cae6da8bc45e3c635d7857b	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-09 15:16:11.177
xnvzmzvtb1i3xywixa350hl2	user_admin_001	admin	update_subject_comment_format	subject	DT	c5b8d1b6a8df7efac8311734:398a6dd68dd393b43a7ff72e49d09628:4fcf32635be34db81579fceb2082b956e251d100df43b5a3f7512eda527ee697cc24cef392b552e05211fa47c8282e339bd2ce9d068e574b28152bbfbbe9639ae7261e700f1e4fbccecebd4ab9804f661ef9781906ca923a60c56e7f5bca6a2b373db1171534abc54085d07c9299cf9550e3048228a94084d78bf946f335f5c885b58966875e1841f35307f58fa0ace96ca5e8c3c66bd47127ec066a5e4ec5328a7389f8b25bf0eca5c1452005fa4e5e786a67c89817	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-08 15:26:49.638
vafrip8rrs5oovd6hnslzbvk	user_admin_001	admin	update_comment_format_template	app_setting	comment_format_template	e313b9727fc5b39496e43924:ba13267bb444a0536c43b57352134ba4:a0c1bffc9879ce7d0b80e0b6fb71594524fb8ccfdf8acf3ea4cfdfad1edc6e186047df6140d59288f31332a7d0af6d5d7f1afd332b2d9ed35f3da88b468fb46dfde8d3a318a9429f12e1ac094209dff084516da7a14fd94dd4a021882e77170caf503ca5a4c095ac12fb0560eac5	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-09 10:45:34.572
dh2w81otxfoo37zj47mh6tqb	user_admin_001	admin	create_class	class	hb9ymtr8xax3ygqr0kdbo874	a77d6125ab77781ffa3cda32:9fea9eb93c91906cd4c1b51dbff66db6:18ef3565bae4957f9c1bfb7221336c048ff58e94f3493607df6e1429bddc10a48d1e3a6ac8b901602fc36821e2ed6a7e4fb819074ce9ec01d3a94c5d3ed2caf9406b2e533702c74d28b2ab7cf68d19d86c64934a3bcc1b7a3d0b8b624b0aa711240428	::1	Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36	2026-05-09 15:16:27.042
\.


--
-- Data for Name: Class; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."Class" ("id", "name", "year", "subjectId") FROM stdin;
528f7d8434d6f3a6f7ad	10Ar/Ar	10	c8c86430aabc681813e8
5315b384ce058e8ab8a6	10Bi1/Bi	10	c5603dd1210b371db7ae
283a8218e792eb7044a2	10Bi2/Bi	10	c5603dd1210b371db7ae
e48d91320929058cd529	10Bs1/Bs	10	4181c022d9eb05117c6d
8082e012c8ef47aeca2d	10Bs2/Bs	10	4181c022d9eb05117c6d
bd14ad4dd640660f1536	10Ch1/Ch	10	b148232e8ddf91bce214
fc3cc49d60e02ea77ee2	10Ch2/Ch	10	b148232e8ddf91bce214
0e647084d6478b6ffe86	10Cs/Cs	10	c8d7bfa257934576575e
3f5084d55890068aef0f	10Ds1/Ds	10	bf43ccb47927cefba4e8
c44e345dd4e716d24329	10Ds2/Ds	10	bf43ccb47927cefba4e8
fda056455d00bf047c21	10Ds3/Ds	10	bf43ccb47927cefba4e8
159a5eac4538c61ab219	10Ds4/Ds	10	bf43ccb47927cefba4e8
848c20d72e2e199b8639	10Dt2/Dt	10	DT
6d7d4c292566a7ad40ef	10En1/En	10	dfa5c7856b40dd6c8c80
24f8f79de7d8daf5cdb1	10En2/En	10	dfa5c7856b40dd6c8c80
69a272fe2ae16f6f7a3d	10En3/En	10	dfa5c7856b40dd6c8c80
4b02a403546e54cc7f5b	10En4/En	10	dfa5c7856b40dd6c8c80
62216e3c0b97d44041eb	10Gcz/Gcz	10	d4b8fb5671d44379b6ca
d0011180a38970fb74e7	10Ge/Ge	10	d0c22c9fbde4eae78a2c
51b76e207ca154e8acba	10IT/IT	10	6d30ecf1c96123a08f97
3e4a03a8fd0bb2afc961	10Ma1/Ma	10	bd0c9a5ae8c3c14861e7
c5a5a84cddf882001a60	10Ma2/Ma	10	bd0c9a5ae8c3c14861e7
dd1b0f7f6f3aefb5955c	10Ma3/Ma	10	bd0c9a5ae8c3c14861e7
d7e3398ffa52ca717c5c	10Ma4/Ma	10	bd0c9a5ae8c3c14861e7
ebee19666521dd14842b	10Mu/Mu	10	d3ca3143b0f1c932737b
d6bb5024d5ec592d3c44	10PE/P1	10	8d14ff7e9d97595155c0
c26810bb7ffda87028cd	10Ph1/Ph	10	450588febb5ff75e8d45
2e959c074f854901899e	10Ph2/Ph	10	450588febb5ff75e8d45
b4234235efeea20a0ca7	10Ps1/Ps	10	0740b4645b4357d8dfed
fd0e3630e13db624ea01	10Ps2/Ps	10	0740b4645b4357d8dfed
a0eea0aac26bc2c9a33b	10a/Re1	10	888ccb16fd8e7504cc67
ec74201ba1740e7d0f91	10ac/Pe1	10	194bd4628e9e1172ea99
24076f2e2af052db9360	10ac/Pe2	10	194bd4628e9e1172ea99
5f5846cdcdc67e6b6535	10b/Re1	10	888ccb16fd8e7504cc67
81ceff273e6f7f288ad5	10c/Re1	10	888ccb16fd8e7504cc67
d8a47a56012378b52d48	11Ab1/Ab1	11	b48a4d58a4743bd8b2db
c6854f08ccb4b2c8da4d	11Ar/Ar	11	c8c86430aabc681813e8
98cc7263651f31776990	11Bi1/Bi	11	c5603dd1210b371db7ae
dfb67ad547da1c74dec0	11Bi2/Bi	11	c5603dd1210b371db7ae
366c6f405b1f8bc78439	11Bs1/Bs	11	4181c022d9eb05117c6d
cbbb379c9906538aaaea	11Bs2/Bs	11	4181c022d9eb05117c6d
8ce3f350dfe12ca44f95	11Ch1/Ch	11	b148232e8ddf91bce214
08cda36fbe812c092c9c	11Ch2/Ch	11	b148232e8ddf91bce214
53889c63251a67f47a2c	11Cs/Cs	11	c8d7bfa257934576575e
4b963bec52944fc95745	11Ds1/Ds	11	bf43ccb47927cefba4e8
6189d54b165dab8bd56e	11Ds2/Ds	11	bf43ccb47927cefba4e8
a82276e194b1732711b4	11Ds3/Ds	11	bf43ccb47927cefba4e8
af2892f8639e42fdc83d	11Ds4/Ds	11	bf43ccb47927cefba4e8
02938d7b62625b948141	11Dt/Dt	11	DT
f27cf82c1214ada6fb4e	11Ec/Ec	11	dfe36bbe8755cfb5ba25
d131a3b33401a7b4a668	11En1/En	11	dfa5c7856b40dd6c8c80
ecf47a0bf5f29be17314	11En3/En	11	dfa5c7856b40dd6c8c80
1b29574d86db3b63e11e	11En4/En	11	dfa5c7856b40dd6c8c80
0bfa86e7a44f47e58019	11En5/En	11	dfa5c7856b40dd6c8c80
d13bbcf819ba7aedcab8	11Fr/Fr	11	45c516b9a78910b2536b
4dbde706728519fe11b9	11Ge/Ge	11	d0c22c9fbde4eae78a2c
e295ccde749fd2982eeb	11Hi/Hi	11	f2b13fe00258ff012e2d
ffc06f66d5b1395c4fe6	11IT/IT	11	6d30ecf1c96123a08f97
96c9cb3e39287f99c251	11Ma1/Ma	11	bd0c9a5ae8c3c14861e7
9035d56968d798f556cb	11Ma2/Ma	11	bd0c9a5ae8c3c14861e7
ff8a7abaa5467496e924	11Ma3/Ma	11	bd0c9a5ae8c3c14861e7
ede2ca1d19344a1b8138	11Ma4/Ma	11	bd0c9a5ae8c3c14861e7
1ef93a4c890e726a2aca	11Mu/Mu	11	d3ca3143b0f1c932737b
15341f322b282b495094	11PE/P1	11	8d14ff7e9d97595155c0
2a8c35f329bd530acc98	11PE1/Pe	11	194bd4628e9e1172ea99
7788942930b1d0ff75cf	11PE2/Pe	11	194bd4628e9e1172ea99
8ecde88989703bbe399b	11Ph1/Ph	11	450588febb5ff75e8d45
f212ea07533e401eb4df	11Ph2/Ph	11	450588febb5ff75e8d45
591c7cad2de417c3536e	11Ps1/Ps	11	0740b4645b4357d8dfed
9ea43fbf389f497e0a95	11Ps2/Ps	11	0740b4645b4357d8dfed
f3f87a4f6848eef065be	11So/So	11	50c923e9e7b41a535325
fbd0f00af5ce74d958eb	11a/Re1	11	888ccb16fd8e7504cc67
a89ce690562f1bfb0bba	11b/Re1	11	888ccb16fd8e7504cc67
6e57422643f6d74880c0	11c/Re1	11	888ccb16fd8e7504cc67
0369837b7f5d361afb51	12Ar/Ar	12	c8c86430aabc681813e8
d664df57dc8d387d1886	12BBU/BBU	12	0f4de0c5ad45017e88f5
50a78e0f85245df78ebb	12Bi/Bi	12	c5603dd1210b371db7ae
2399822855f02dd17bd4	12Bi2/Bi2	12	c5603dd1210b371db7ae
f30291eeab3f00e3032f	12Bs/Bs	12	4181c022d9eb05117c6d
7c1a4815aa1a2d69e547	12Ch/Ch	12	b148232e8ddf91bce214
46d83744c63768dc907e	12Ch2/Ch	12	b148232e8ddf91bce214
0b72791b2dd715141534	12Dm/Dm	12	abff90061fc53561024e
29e5c9ab72aa19cf5199	12EN/En	12	dfa5c7856b40dd6c8c80
242ae56440681672b491	12FM1/FM	12	34b041927ed21b2b5cae
af736540ce33e44663db	12FM2/FM	12	34b041927ed21b2b5cae
86f346656a8fafa3495b	12Ma1/Ma	12	bd0c9a5ae8c3c14861e7
ebfcbb48fbd63e18339e	12Ma2/Ma	12	bd0c9a5ae8c3c14861e7
7f1e9e42b1894464421d	12Ma3/Ma	12	bd0c9a5ae8c3c14861e7
a5012b7bbf099fe29fe8	12Me/Me	12	4096cb88b43e598d2a72
f6dcd719304f00d82c2e	12Ph1/Ph	12	450588febb5ff75e8d45
bb68063696175471a537	12Ps/Ps	12	0740b4645b4357d8dfed
a72e2dd120ef2c86ef04	12Pt/Pt	12	8fbfd22f3e60c619e941
7e2d798388bf048f20f6	12So/So	12	50c923e9e7b41a535325
1b092bfce595f85293ee	12St/St	12	b1dd6ef556c9ec20224b
b220f2ecaea0a0fb00dc	12a/Re1	12	888ccb16fd8e7504cc67
6e2aeb5d8592447a31d6	12b/Re1	12	888ccb16fd8e7504cc67
eb6a091792f47c508880	13Bi/Bi	13	c5603dd1210b371db7ae
c3c9a4fcec3d5a4cc8f5	13Bs/Bs	13	4181c022d9eb05117c6d
c7844f6432c9b3505784	13Ch/Ch	13	b148232e8ddf91bce214
e6e6b9eaef349d8f67bb	13DM/Dm	13	abff90061fc53561024e
5d0f2a46fc3857533f25	13En1/En	13	dfa5c7856b40dd6c8c80
b29c20c00222a6c1876c	13Fr1/Fr	13	45c516b9a78910b2536b
75a0ab6961da49ab3b89	13Ge/Ge	13	d0c22c9fbde4eae78a2c
2aba875ca65d820f85b4	13Ma1/Ma	13	bd0c9a5ae8c3c14861e7
390155dc1ec0a20487bd	13Ma2/Ma	13	bd0c9a5ae8c3c14861e7
30720b1ba7b7ea205543	13Ph/Ph	13	450588febb5ff75e8d45
00a45de906cbd0942534	13Ps/Ps	13	0740b4645b4357d8dfed
f75e70904a70afafbbd1	13Pt/Pt	13	8fbfd22f3e60c619e941
552ab552eb8762c570a0	13St/St	13	b1dd6ef556c9ec20224b
78a0f435af1bd5acb2c7	13a/Re1	13	888ccb16fd8e7504cc67
9bbe8ed0f5c6182ff324	13b/Re1	13	888ccb16fd8e7504cc67
6f69f172e874d825216c	7a/Ab1	7	b48a4d58a4743bd8b2db
40e35d563990a350b589	7a/Ab2	7	8c692847d52879b4be83
ade303ded2be972c6a5d	7a/Ar1	7	c8c86430aabc681813e8
37467e477610598c1474	7a/Dr1	7	604ab32c3290fb9c3a31
ce89e455bbdc905dd759	7a/Dt1	7	DT
5e6582e3fe3b905468a7	7a/Fr1	7	45c516b9a78910b2536b
98559a16949b4230e009	7a/Ge1	7	d0c22c9fbde4eae78a2c
6b755b3363797e6feca4	7a/Hi1	7	f2b13fe00258ff012e2d
d13ffb488e82f76b02d0	7a/It1	7	6d30ecf1c96123a08f97
310c982ca10f2d32ccb1	7a/Mu1	7	d3ca3143b0f1c932737b
d421b774241434ccd836	7a/Pe1	7	194bd4628e9e1172ea99
39b1ab8e2df8c0182786	7a/Re1	7	888ccb16fd8e7504cc67
b630f79c61b93bb9a72b	7a/Sc1	7	2d060f26699b06743f31
3649e9e9359e028f0892	7a/Sp1	7	ff1729d8c7e693c5b111
b6bbbaae703037f495be	7a/pshe	7	647cdbe0df378408fdef
0aeb1b79309469c06c74	7ac/En1	7	dfa5c7856b40dd6c8c80
567b97f8b1de47a17d7d	7ac/En2	7	dfa5c7856b40dd6c8c80
a7de483e1b81ba1ede04	7ac/En3	7	dfa5c7856b40dd6c8c80
0dae9224680dbe8d6673	7ac/Ma1	7	bd0c9a5ae8c3c14861e7
492be8b28f568157cbef	7ac/Ma2	7	bd0c9a5ae8c3c14861e7
d3f8ea02cba4d989b5f4	7ac/Ma3	7	bd0c9a5ae8c3c14861e7
89788763a94d40f933fc	7ac/Ma4	7	bd0c9a5ae8c3c14861e7
94b4fdf386152f607a13	7ad/En4	7	dfa5c7856b40dd6c8c80
0f69da4484eb0767a720	7b/Ab1	7	b48a4d58a4743bd8b2db
dc667f3c1235dee04b67	7b/Ab2	7	8c692847d52879b4be83
0e8355bb2b56b6ccda79	7b/Ar1	7	c8c86430aabc681813e8
96ead37b6d717ecf4c9d	7b/Dr1	7	604ab32c3290fb9c3a31
dd84dd43e18cd13b299e	7b/Dt1	7	DT
7f4d00c651d1bc22f95e	7b/Fr1	7	45c516b9a78910b2536b
be3719af94743bcb4e1b	7b/Ge1	7	d0c22c9fbde4eae78a2c
c34e64cafe735b50ed2e	7b/Hi1	7	f2b13fe00258ff012e2d
c69b3760d1f706da6487	7b/It1	7	6d30ecf1c96123a08f97
f82c00f0fb8ba676d79a	7b/Mu1	7	d3ca3143b0f1c932737b
d5cde185f3020a8d2c79	7b/Pe1	7	194bd4628e9e1172ea99
da2af99abb1da9ed1063	7b/Re1	7	888ccb16fd8e7504cc67
ac8714c2edc7a6940ce7	7b/Sc1	7	2d060f26699b06743f31
335b5ae685bab3b50f30	7b/Sp1	7	ff1729d8c7e693c5b111
56061cc582e767291018	7b/pshe	7	647cdbe0df378408fdef
0ab34b12507bd5ae0386	7c/Ab1	7	b48a4d58a4743bd8b2db
09c52f1ac992e352a953	7c/Ab2	7	8c692847d52879b4be83
d15d7b9fd2ed7e46dd53	7c/Ar1	7	c8c86430aabc681813e8
859e00b64f3807defcfe	7c/Dr1	7	604ab32c3290fb9c3a31
87ccbe2ce9a680ff93a0	7c/Dt1	7	DT
a477427b675c489d36f6	7c/Fr1	7	45c516b9a78910b2536b
0e634788a37d82abbef4	7c/Ge1	7	d0c22c9fbde4eae78a2c
d7cef190d3651254a780	7c/Hi1	7	f2b13fe00258ff012e2d
2a9db3e19b93ce9c355d	7c/It1	7	6d30ecf1c96123a08f97
98d9b41d2eb6d4fc61c8	7c/Mu1	7	d3ca3143b0f1c932737b
dbe35887465186683743	7c/Pe1	7	194bd4628e9e1172ea99
7e1c155a4ca800a30d86	7c/Re1	7	888ccb16fd8e7504cc67
8cdd523b41f1a026f4de	7c/Sc1	7	2d060f26699b06743f31
061ddff82e7b822b7ca7	7c/Sp1	7	ff1729d8c7e693c5b111
abe9d4482e5fa2c5fa0f	7c/pshe	7	647cdbe0df378408fdef
05504bdd46170dd3f4db	7d/Ab1	7	b48a4d58a4743bd8b2db
83e93fed391058692f49	7d/Ab2	7	8c692847d52879b4be83
6db953d26eebb36848ad	7d/Ar1	7	c8c86430aabc681813e8
22eceb505d094f53c53e	7d/Dr1	7	604ab32c3290fb9c3a31
a8ff0afc8a2eaa57d68b	7d/Dt1	7	DT
273f5b438cc3414ba44f	7d/Fr1	7	45c516b9a78910b2536b
ad8ed2ff59798b5df7ba	7d/Ge1	7	d0c22c9fbde4eae78a2c
8639aeda2bf449728d9d	7d/Hi1	7	f2b13fe00258ff012e2d
b4ccb16d04bc3b14647d	7d/It1	7	6d30ecf1c96123a08f97
b4f371ad7826b1be4b8f	7d/Mu1	7	d3ca3143b0f1c932737b
bc02a71806dababec353	7d/Pe1	7	194bd4628e9e1172ea99
4d66ba7137d37d690550	7d/Re1	7	888ccb16fd8e7504cc67
d9c04b00fffcf9b34420	7d/Sc1	7	2d060f26699b06743f31
011da44732c7f182de66	7d/Sp1	7	ff1729d8c7e693c5b111
0af02ffbcd58d7253c76	7d/pshe	7	647cdbe0df378408fdef
15d8cb818f560a27d9a8	8a/Ab1	8	b48a4d58a4743bd8b2db
ccd8ac5ce51d687a65b6	8a/Ab2	8	8c692847d52879b4be83
ccdd93e22ce8fde4aefa	8a/Ar1	8	c8c86430aabc681813e8
81f8ac02e568df5d9552	8a/Ck1	8	6196f2f5279ed1941ab7
3ea70d2daf15198c70a6	8a/Dr1	8	604ab32c3290fb9c3a31
4a894a4948ebc3d652ed	8a/Dt1	8	DT
62a6aa3edf4fe2dee2d0	8a/En1	8	dfa5c7856b40dd6c8c80
cd4dde106f1d53c9aada	8a/Fr1	8	45c516b9a78910b2536b
011f0cc85f1ad1cc22e6	8a/Ge1	8	d0c22c9fbde4eae78a2c
be0802e0a01e96e75eab	8a/Hi1	8	f2b13fe00258ff012e2d
2ee37f6a5b5afdba3728	8a/It1	8	6d30ecf1c96123a08f97
042f404c813aed6ee038	8a/Mu1	8	d3ca3143b0f1c932737b
9538a1857d657c0f004b	8a/Pe1	8	194bd4628e9e1172ea99
998e79a5a2df8f596fa3	8a/Re1	8	888ccb16fd8e7504cc67
bd6f2ca2ed437f9499d2	8a/Sc1	8	2d060f26699b06743f31
d74ad31e1cd4b3e1a678	8a/pshe	8	647cdbe0df378408fdef
9b82d77926477514bec2	8ac/Ma1	8	bd0c9a5ae8c3c14861e7
dcfa4362bb9ba014e5b6	8ac/Ma2	8	bd0c9a5ae8c3c14861e7
70d175493126d8dceaf2	8ac/Ma3	8	bd0c9a5ae8c3c14861e7
67cb84d4c26e5522b07c	8ad/Ma4	8	bd0c9a5ae8c3c14861e7
7d66aa886cfe3a72f86b	8b/Ab1	8	b48a4d58a4743bd8b2db
d8bb7e8a8ebe34b835c3	8b/Ab2	8	8c692847d52879b4be83
1993cf9763443fce0f82	8b/Ar1	8	c8c86430aabc681813e8
4901021e0e8f2486bbf9	8b/Ck1	8	6196f2f5279ed1941ab7
b1708239145b8854c8a3	8b/Dr1	8	604ab32c3290fb9c3a31
17be688be816f0468b05	8b/Dt1	8	DT
6e3f730a3b6e866694e3	8b/En2	8	dfa5c7856b40dd6c8c80
7dc8596f01b7d5f27d2f	8b/Fr1	8	45c516b9a78910b2536b
58f247c4f27334049744	8b/Ge1	8	d0c22c9fbde4eae78a2c
417a1d3104a891ccfa38	8b/Hi1	8	f2b13fe00258ff012e2d
9a01025dd44aaf5de999	8b/It1	8	6d30ecf1c96123a08f97
fc618b2ec9d2df99322d	8b/Mu1	8	d3ca3143b0f1c932737b
52a0c7618351d4fe83e9	8b/Pe1	8	194bd4628e9e1172ea99
332ec0589e88163cdbc2	8b/Re1	8	888ccb16fd8e7504cc67
c831500edcf62a572d96	8b/Sc1	8	2d060f26699b06743f31
942711e71cb8fb753a49	8b/pshe	8	647cdbe0df378408fdef
140cb8e7fd7f918244a8	8c/Ab1	8	b48a4d58a4743bd8b2db
0f9cf77856a913b30c8e	8c/Ab2	8	8c692847d52879b4be83
67a7363447e757702674	8c/Ar1	8	c8c86430aabc681813e8
a09b6ca1c047b268f7cc	8c/Ck1	8	6196f2f5279ed1941ab7
ffecd66e5100127c4042	8c/Dr1	8	604ab32c3290fb9c3a31
6c05b760b3d8e03e72a7	8c/Dt1	8	DT
937d339e9ff57e16d3c9	8c/En3	8	dfa5c7856b40dd6c8c80
dac65aeaaed135b88dba	8c/Fr1	8	45c516b9a78910b2536b
4883160b03bc92b91640	8c/Ge1	8	d0c22c9fbde4eae78a2c
d87f1150ff27c1efbcc2	8c/Hi1	8	f2b13fe00258ff012e2d
524cb475f7eec63c3736	8c/It1	8	6d30ecf1c96123a08f97
9b8fac85cf8101939f27	8c/Mu1	8	d3ca3143b0f1c932737b
0716aa5bb459f01a8e41	8c/Pe1	8	194bd4628e9e1172ea99
fe2dc648d4dc41df5d30	8c/Re1	8	888ccb16fd8e7504cc67
42c2357c9e79a8d2bd9e	8c/Sc1	8	2d060f26699b06743f31
1bfba56a05f5800f2691	8c/pshe	8	647cdbe0df378408fdef
c38a7595c8fe93bb3a4e	8d/AB1	8	b48a4d58a4743bd8b2db
8b094bf3bca9e805dbab	8d/AB2	8	8c692847d52879b4be83
922a219336536e10a18d	8d/Ar1	8	c8c86430aabc681813e8
5e3639a7c71b3a997aa8	8d/Ck1	8	6196f2f5279ed1941ab7
53fd09a399c9afcd820e	8d/Dr1	8	604ab32c3290fb9c3a31
ad93cb9499239edb202a	8d/Dt1	8	DT
1c95850b2ece593eeb02	8d/En4	8	dfa5c7856b40dd6c8c80
0da1a8e4012fbe1989cc	8d/Fr1	8	45c516b9a78910b2536b
83db990ac81dfd8b35c7	8d/Ge1	8	d0c22c9fbde4eae78a2c
74b1f4902d76a7737fad	8d/Hi1	8	f2b13fe00258ff012e2d
7a6c12bdaf5c6da7b21d	8d/It1	8	6d30ecf1c96123a08f97
6128f2c7422b743e9a52	8d/Mu1	8	d3ca3143b0f1c932737b
26f71e77ebb5d7b35f70	8d/Pe1	8	194bd4628e9e1172ea99
8899967fd022bd010f96	8d/Re1	8	888ccb16fd8e7504cc67
f11f70d30bce575096ec	8d/Sc1	8	2d060f26699b06743f31
3810d251d2a035b1de85	8d/pshe	8	647cdbe0df378408fdef
457e892a683f70a78c8a	9a/Ab1	9	b48a4d58a4743bd8b2db
9a64dc0d5caa2750bf59	9a/Ab2	9	8c692847d52879b4be83
471952423957373168f7	9a/Ar1	9	c8c86430aabc681813e8
0b701012c2790c7c4735	9a/Dt1	9	DT
1f2141b07765892cb67a	9a/ENT1	9	5eca2fa95cdf368d7fba
91df72a6767bb7c3ab87	9a/En1	9	dfa5c7856b40dd6c8c80
7e5ffa821f50c0bc8f8a	9a/Fr1	9	45c516b9a78910b2536b
862d7a9266b086d90ae6	9a/Ge1	9	d0c22c9fbde4eae78a2c
117e47621ce29f1a2bc4	9a/Hi1	9	f2b13fe00258ff012e2d
0371cd4bdfd4b24f8cbe	9a/It1	9	6d30ecf1c96123a08f97
877d41f9b7669c490ac5	9a/Mu1	9	d3ca3143b0f1c932737b
97d9dad9b451c17681bc	9a/Re1	9	888ccb16fd8e7504cc67
f37759490150fb4b8c33	9a/pshe	9	647cdbe0df378408fdef
b8232c40c37937eb01d7	9ab/Pe1	9	194bd4628e9e1172ea99
ed07ec8fa564381ea093	9ab/Pe2	9	194bd4628e9e1172ea99
3c98fafabb219e9ecd4d	9ac/Ma1	9	bd0c9a5ae8c3c14861e7
b95cadce5af0089d8b98	9ac/Ma2	9	bd0c9a5ae8c3c14861e7
b07f3d9674f357d4fd0d	9ac/Ma3	9	bd0c9a5ae8c3c14861e7
faf6df81a0493135c916	9ac/Ma4	9	bd0c9a5ae8c3c14861e7
67b3b7100d976861d8b7	9ac/Sc1	9	2d060f26699b06743f31
2bd14e1ccaaa323968fe	9ac/Sc2	9	2d060f26699b06743f31
3d2845b4e39b80d24680	9ac/Sc3	9	2d060f26699b06743f31
b96a13884d9cb2f18ad3	9ac/Sc4	9	2d060f26699b06743f31
99e6504888bd1d9f9986	9b/Ab1	9	b48a4d58a4743bd8b2db
c90d77745ab6d0e48537	9b/Ab2	9	8c692847d52879b4be83
bd6be185e5e6e11ada1b	9b/Ar1	9	c8c86430aabc681813e8
2f7f607a87c99c6722e9	9b/Dt1	9	DT
dd1d596229411ff299e5	9b/ENT1	9	5eca2fa95cdf368d7fba
9bb832cf61494f6878ed	9b/En2	9	dfa5c7856b40dd6c8c80
04dd4b67bcf35901231b	9b/Fr1	9	45c516b9a78910b2536b
74f2bc4a5877a6277d7c	9b/Ge1	9	d0c22c9fbde4eae78a2c
649153b9a1951bb4bb9a	9b/Hi1	9	f2b13fe00258ff012e2d
76d5185f429653ed2e87	9b/It1	9	6d30ecf1c96123a08f97
bb3d38b9bf77a8e95b79	9b/Mu1	9	d3ca3143b0f1c932737b
431824f1591821900428	9b/Re1	9	888ccb16fd8e7504cc67
dd3c19d41c8dd7d27682	9b/pshe	9	647cdbe0df378408fdef
8427cb0e609c6395011d	9c/Ab1	9	b48a4d58a4743bd8b2db
13cf5f07e77de99dcc30	9c/Ab2	9	8c692847d52879b4be83
cf5177242c1034802d8e	9c/Ar1	9	c8c86430aabc681813e8
7f18686d985523c66aa3	9c/Dt1	9	DT
b232a97de45a99e40e65	9c/ENT1	9	5eca2fa95cdf368d7fba
5c960d1023e68e8b8d68	9c/En3	9	dfa5c7856b40dd6c8c80
50539e1ab2fe53f5df15	9c/Fr1	9	45c516b9a78910b2536b
86c46654bedeb5a02591	9c/Ge1	9	d0c22c9fbde4eae78a2c
0f692345829560ccc163	9c/Hi1	9	f2b13fe00258ff012e2d
ec964e12ec298f154906	9c/It1	9	6d30ecf1c96123a08f97
e48a2537426fd384e04b	9c/Mu1	9	d3ca3143b0f1c932737b
028f24b1257fd6f8b5af	9c/Re1	9	888ccb16fd8e7504cc67
0a30a4d85b028245ef62	9c/pshe	9	647cdbe0df378408fdef
3ce1e9bfc95928b110dc	9cd/Pe1	9	194bd4628e9e1172ea99
c5aa80e1a22c4b09fbfe	9cd/Pe2	9	194bd4628e9e1172ea99
4d780077a193e1d55df4	9d/Ab1	9	b48a4d58a4743bd8b2db
809db13c3bb696a97ee7	9d/Ab2	9	8c692847d52879b4be83
1094e4c9fe718bcb6290	9d/Ar1	9	c8c86430aabc681813e8
ab845c01757acc0d395d	9d/Dt1	9	DT
b2e5e7431c984ac1c0d7	9d/ENT1	9	5eca2fa95cdf368d7fba
f6eeb477d5f1c37726c6	9d/En4	9	dfa5c7856b40dd6c8c80
81f20f3be795fb6d955d	9d/Fr1	9	45c516b9a78910b2536b
9e72e29c3c0b838bfc4c	9d/Ge1	9	d0c22c9fbde4eae78a2c
19cc8e8396329ca6c0bc	9d/Hi1	9	f2b13fe00258ff012e2d
65d085a17aba3a9b1d94	9d/It1	9	6d30ecf1c96123a08f97
5d13ab8d40255bc00b5e	9d/Mu1	9	d3ca3143b0f1c932737b
acb5f26d24423de54ac7	9d/Re	9	888ccb16fd8e7504cc67
4e4f54ae37e0e224fdf3	9d/pshe	9	647cdbe0df378408fdef
\.


--
-- Data for Name: CommentGroup; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."CommentGroup" ("id", "name", "displayOrder", "subjectId", "title", "isLinked", "linkedField") FROM stdin;
sdt5wca5bot6d8lo2u3n61yx	Communication	1	DT	Communication of Design Ideas	f	\N
md1b4mj0f6ztl29ksjkani72	Make Skills	2	DT	Manufacturing Skills	f	\N
yb8henvd4z79t4xljb1ciu3q	Materials And processes	3	DT	Materials and Processes Knowledge	f	\N
xdtpcx4ay2u3o0mm9fuuso1l	Units Studied	0	DT	Units Studied	f	\N
\.


--
-- Data for Name: CommentOption; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."CommentOption" ("id", "code", "text", "displayOrder", "groupId") FROM stdin;
bymqwmprlgflgl5ps40bjwdm	7	This year, pupils were introduced to Design and Technology, developing safe working practices and drawing skills through Isometric Sketching. They applied these to CAD/CAM projects, manufacturing 2.5D Door Hangers and British Clocks, before studying Sustainability and the 6 Rs.	1	xdtpcx4ay2u3o0mm9fuuso1l
ll88auca6utxs8je6ivdds4k	8	This year, pupils revisited Design and Technology foundations and reinforced safe working practices. They extended skills into three dimensions through the 3D candleholders project and explored Sustainability — Energy Generation.	2	xdtpcx4ay2u3o0mm9fuuso1l
kzzm8ki6qgg2n4f50jadnmst	9	This year, pupils refreshed their understanding of Design and Technology and reaffirmed safe working practices. They studied mechanisms in the Cams and Gears unit and completed a Sustainability Design — Textile Bags project.	3	xdtpcx4ay2u3o0mm9fuuso1l
f1x5gjx4qu7q1d9n3t93pseb	High	<Name> communicates <his> design ideas in <Subject> with exceptional clarity. <His> sketches are detailed and thoroughly annotated, using rendering and orthographic projection to convey form and function.	1	sdt5wca5bot6d8lo2u3n61yx
qaz0yalxourn46tzosgcax4s	Med	<Name> communicates <his> design ideas clearly in <Subject> using annotated sketches. To progress, <Name> should deepen <his> annotation — explaining why design choices are made, not just what they are.	2	sdt5wca5bot6d8lo2u3n61yx
t5hsops220j5bu4sm3ikmoc3	Low	<Name> needs to develop how <he> communicates design ideas in <Subject>. <His> sketches lack detail and annotation. <He> should practise drawing techniques and use prompts to explain <his> thinking.	3	sdt5wca5bot6d8lo2u3n61yx
x59xbal93acd84xoikhifcu3	High	<Name> demonstrates excellent practical skill in the <Subject> workshop, independently selecting tools to produce high-quality outcomes. <He> evaluates <his> own making and adjusts <his> approach to improve results.	1	md1b4mj0f6ztl29ksjkani72
h77f394372uiiqeppzsriyfi	Medium	<Name> works safely and confidently in <Subject>, producing generally accurate work. To progress, <he> should slow down at key stages — marking out, measuring and finishing — to improve the final quality.	2	md1b4mj0f6ztl29ksjkani72
owrb9f5mnp5gp77sjgu6d18t	Low	<Name> requires regular support to work safely and accurately in the <Subject> workshop. <He> struggles to follow procedures in order. <Name> should listen to demonstrations and check measurements before cutting.	3	md1b4mj0f6ztl29ksjkani72
s3m64yzzb2xbsic8u9e6md4h	High	<Name> shows a sophisticated understanding of materials and processes in <Subject>. <He> confidently classifies materials, describes their properties and explains manufacturing processes, consistently justifying <his> choices with technical reasoning linked to user needs, function and sustainability.	1	yb8henvd4z79t4xljb1ciu3q
upb1bj9odzuy8djekhexwo6e	Medium	<Name> has a sound knowledge of materials and processes in <Subject>, naming materials, describing their properties and explaining familiar manufacturing steps. To progress, <he> should use more precise technical vocabulary and justify <his> choices.	2	yb8henvd4z79t4xljb1ciu3q
kasbyby331o69w6kusa1u88f	Low	<Name> has a limited grasp of materials and processes in <Subject> but will improve with time. <He> should use knowledge organisers to recall terminology and describe common manufacturing processes.	3	yb8henvd4z79t4xljb1ciu3q
\.


--
-- Data for Name: CommonCommentGroup; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."CommonCommentGroup" ("id", "name", "title", "displayOrder", "isLinked", "linkedField") FROM stdin;
m7k11htb7ybhjbzra6a0pbx8	Academic	Academic Performance	1	f	\N
gq2e46phsivlcljdybugoaw6	Behaviour	Behaviour	2	f	\N
fht0kqpvp21kxsfwq6hyalbm	SocialSkills	Social Skills	3	f	\N
n5sy0axnb4oj9k0nvnyxkc3o	Collaboration	Collaboration	4	f	\N
ismry4x2kzdwssn2pdbkbu8f	Overall	Overall	6	f	\N
kvyc23qpstalii62z2ytnrl5	Units Studied	Units Studied	0	f	\N
\.


--
-- Data for Name: CommonCommentOption; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."CommonCommentOption" ("id", "code", "text", "displayOrder", "groupId") FROM stdin;
r3rga6rgas2akmt7m6i4ktok	7	Not Set	0	kvyc23qpstalii62z2ytnrl5
ga18zj0kz662x618tabt6813	8	Not Set	1	kvyc23qpstalii62z2ytnrl5
t3n6ux6qdio6uxkigk8y67x5	9	Not Set	2	kvyc23qpstalii62z2ytnrl5
tws0wzviep9dhs872nqidd89	High	<Name> is a model of excellent behaviour for learning. <He> arrives fully equipped, listens carefully and engages thoughtfully with every task. <Name> is courteous to staff and peers, consistently making the most of <his> learning.	0	gq2e46phsivlcljdybugoaw6
ykihn6of36l5fylrwixku2ty	Medium	<Name>'s behaviour is generally good. <He> follows instructions and is respectful to staff and peers, though <he> can be distracted at times and occasionally needs reminding to refocus.	1	gq2e46phsivlcljdybugoaw6
ghklrglxe1szmginmmq4mycf	Low	<Name>'s behaviour is a barrier to <his> progress and others around <him>. <He> requires frequent reminders to follow instructions and stay on task. <Name> must take greater responsibility for <his> choices in class.	2	gq2e46phsivlcljdybugoaw6
up91qk5epq44y6x6apw3x7p5	High	<Name> demonstrates excellent social skills. <He> communicates politely and confidently with staff and peers, shows genuine empathy and handles disagreements calmly, having a positive effect on the classroom atmosphere.	0	fht0kqpvp21kxsfwq6hyalbm
n5nkfa0ysm5wc4kwwlco5hm6	Medium	<Name>'s social skills are developing well. <He> is generally polite and contributes positively to lessons. <He> should pause to consider the impact of <his> words and actions when managing disagreements.	1	fht0kqpvp21kxsfwq6hyalbm
ma5e59t3bpicl9secnlklsda	Low	<Name>'s social skills need development. <He> can struggle to interact appropriately, sometimes responding impulsively. <Name> should speak respectfully, listen carefully and seek adult support when situations become difficult.	2	fht0kqpvp21kxsfwq6hyalbm
tumuxai35otinyej62mb8ceg	High	<Name> is an outstanding collaborator. <He> works confidently in groups, takes on different roles and ensures every member can contribute. <Name> builds on others' ideas to reach a stronger shared outcome.	0	n5sy0axnb4oj9k0nvnyxkc3o
b88c0a8lneuld4wum8roaxse	Medium	<Name> works productively with others, sharing ideas and supporting teammates. To strengthen <his> teamwork, <Name> should actively involve quieter group members and be willing to compromise when ideas differ.	1	n5sy0axnb4oj9k0nvnyxkc3o
w01iob6b0o2sglgzl4ipepyy	Low	<Name> finds collaborative work challenging, tending to either dominate or disengage entirely. <Name> should focus on listening to others' suggestions, sharing the work fairly and accepting that strong outcomes require compromise.	2	n5sy0axnb4oj9k0nvnyxkc3o
gug2j5s550fp1wq2pg0f8214	High	Overall, <Name> has had an excellent year in <Subject>, approaching lessons with curiosity and effort, reflected in the high quality of <his> work. To flourish further, <he> should seek challenge and reflect on how to improve each outcome.	0	ismry4x2kzdwssn2pdbkbu8f
dyzw23jb3hndcr8o6parcm6u	Medium	Overall, <Name> has made solid progress in <Subject> this year, working sensibly and demonstrating understanding of the core ideas. To reach the upper bands, <Name> should apply greater care to <his> design work, practical finish and technical vocabulary.	1	ismry4x2kzdwssn2pdbkbu8f
bqp96t6yfdbs2lzx6cdnf2u7	Low	Overall, <Name> has found aspects of <Subject> challenging this year. With more consistent focus and a willingness to ask for help, <Name> is capable of real progress. The priority for next year is to take ownership of <his> learning.	2	ismry4x2kzdwssn2pdbkbu8f
cs3r0ad7jwxtv9l4dfvvukce	AP	This year, <Name> has achieved an end of year level of <EoYLevel> against a target grade of <TargetLevel>.	0	m7k11htb7ybhjbzra6a0pbx8
\.


--
-- Data for Name: CommonPupilCode; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."CommonPupilCode" ("id", "assignmentId", "commonGroupId", "code") FROM stdin;
\.


--
-- Data for Name: Deadline; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."Deadline" ("id", "title", "date", "description", "isActive", "createdAt") FROM stdin;
\.


--
-- Data for Name: IgnoredWord; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."IgnoredWord" ("id", "teacherId", "word", "createdAt") FROM stdin;
40e4f613-0063-4636-b680-0210eaa6babd	user_admin_001	organisers	2026-05-03 15:50:29.796727+00
2d59c687-d5d3-4a70-93de-67e4362b5b72	user_admin_001	organisation	2026-05-03 15:51:43.35419+00
\.


--
-- Data for Name: Pupil; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."Pupil" ("admissionNumber", "firstName", "lastName", "gender", "isActive", "form") FROM stdin;
1242	f367838e31b2cb02f0dee0d8:23813d76098055c945f8011a1c80f2df:71313a7891	fe5b604c50be836ed8da70cb:51cfeba140b9db20fe794327e3cb2394:c1aaa51f4f6eb727	F	t	7A
807	d0927ec1e7eec394eb40dcab:231bcabf4a028759992be5d6de4f1774:e22e9e9f	d90ede3906a5c66d4c06645c:676c17414ba11978e07e3bed12c8be1f:c7da4a45c48ffa07e75285	F	t	7A
2287	55fa1913206d40ad0f29bf76:9dbf5f981a002e2b814e6c06ba55f009:5aca6e466fb442	258bbf75632fa393f2fe82fc:a6a2b86acdc43d4953ea8004d6c209c5:178d41af25	M	t	7A
2299	d1e1fc0aceb82210093be328:d560a55b892998d18aac2a5915e3f57f:2b1baffc6e	71fddd041685003fe82c6698:e8b365bf0af7f88dbe2394e3bd95b902:32cd85724996f872894e304ebfe3b558	M	t	7A
2587	9ddf3f61cbd754cb63211ddd:1b05f21ffad9f53ba56e0ca0771578b0:9dd5a95d	6f03624eee3dbeb9d8472acf:62a79828cd72fa62cc17df65246ba913:3e82129e5143fd4f	M	t	7A
1341	69d2c904c2a92d14f134617e:b70026b69e0a45a2371ee693de2ef80c:763bf907a0eb98ca463e2a	08c3d6ecb188ff8cbdcb4aa9:483b6d38317c28ec9992bc9e565e2665:3b79023cbd5341e5	M	t	7A
2560	dfdb4778ef2f8d5a931181a3:8119beb201d346486fcaf58b5bbcbec6:1c2e828d	5e0024c87a2ab8dd7b348eae:8f17c0ebb908ed7738b97718651189df:a82fd137	F	t	7A
811	b56543cfbec5489ae01d0420:68309abd6b991faa0005128f6475f6df:f56fa7227a53d1	7c925ae1226f20371d4e23af:7323c068e6438556caa35f5307eda241:eec1257f2e0524	M	t	7A
2413	3b4416ceb97904f5c69b42a3:dedfcecd21a11c88739580def858dcf8:4334b63cb4	e7952810738dc3aedc90b5ea:141a4b60f67d702ce3196e02ad9b21cc:5888a9c30b15821ba322cd	F	t	7A
1954	4ffbcf9954141e6035046ae4:5b7bd876e3e2d3d93ec5b9cbd0334a2b:bce1b25ed9	958b2c78e2878722898a3fb8:d69034eebff18a3a09db43ab30796550:56eb086aada553	F	t	7A
2381	853345e9c6e563990c84ed39:1e4c6b85a313094a1dc3c1518a4e8343:64849dcae3	c51a038dd4ae901297cc2eeb:b7d064d5be525ceb98c127c182d698cd:1c00f844f472	F	t	7A
2416	28b1333a393c0199e06bae62:0b038ddab471ec274aed05fc35531b54:9882cef2e0	4db2e44270d475855bf7a1f3:10ca5c30dd940219ce35a04b6d891ce5:7b997931e3	M	t	7A
2036	f4439d4589dfb7c90851e316:4d1c22909c015b001f37925b26d7c309:2320b829c35ebaf9	0a903be4e8661e5a38314d8f:26082476df55108043cba25b1661a0e2:e034cdf8505d	M	t	7A
2259	cbad304ceef75466924e857b:25962abefc15f81d0028946489ac701f:de0ed1130a2a	db401c38b4dff762010c8e9b:b8aa40b86ba27c8be067525fd1ca6801:81ef2fd64ac3	M	t	7A
2254	aedbb6b132b01b03b6bda344:1677f3eedbb376eb7dd56a43df0605d4:c5079d91	e39d0c1dd91483e20eb299e8:244a7cf28080e78e67acb3cd43f5719a:b295730c6259	F	t	7A
2331	8c7ef42f460b76da04ab92d3:46ddf1911c104ca7d8f3dd68a10e528e:3e913e505354	a5e4f8464a7bf6b81dbc708b:ce3893acbc48ab4c4a1c6084a42eba76:3d49b8f2e1bc20	F	t	7A
2232	32f5a8b473d3ec074089d0c8:8bbbe2e0ee2b75741bc02f2c9cad2e27:169b847940b442	bf5d52394d7c1e869140dc84:9d27f704f69a52ca66668e9ae89256c4:371800be17ff7929ba	M	t	7A
844	5725d96860238220be3337f3:ef43ce7a3b227411fac6504cbfa7a078:d04c7297387b	ca3b297a4882c4187b09c8f6:1012d6b63aff3f91c4ae4f5981bf198c:06e353d3bb9c	M	t	7A
2195	6ed85de40c9a2dd74a2f796c:d30e906d17fcf9f4caa54cb890b91429:63a46e88f0d7431c	926fb6568e20899296db20b1:e103b5777395e16747a701e0332a98af:c9ddab02	M	t	7A
2135	0b3864fc162bada43b5228ca:d34ff427c161b8efacff8cb49c3b87f8:4968311f99	f98cdfe359068e459daacdaf:ed0caf003f3c6d75bf782573d6f8111d:d912a0352118a8	F	t	7A
847	904782bfca0a9ef129e3cda2:c293d2ba34476bda87288ddeebcdc67d:0b78995c1f15	a42e09506f1250a5b5db2538:8781f5da55534bd46f55f94713dc5177:7144e26b924c	F	t	7A
1964	c71de362a7fbbf5d992457df:b9226d84ebfa0f4d092af112eb875e60:de9f7b793e	e0b9d82d87ce7ddcf6804049:86095010c88fab474fef6918981cb257:e5b9a932f5	M	t	7A
1240	0bc2ad5aa94f9d44093a8b6b:cd838e46cb8330f2292b748db7017314:072171bd55	e9f5cf4cf0bbbbf5de030461:9a119d8008bac0423a47310e907b7217:6e268d5571	F	t	7B
809	978c639635a43dc6b9874cae:e2b68ebcb46a17a3029529ea72c76e0c:bd6faa37	5b5b1b870e3096b3358aab5f:a238c1c0aa3279f4f5b544b4aa7be0c5:529e2facc3eaf2beb0ddb7a6	F	t	7B
1996	cbd7c84606ff5c58fb187f47:38d2720839bcc9c9b92e6567930a57ab:f1537f84dd43	cf6018c0dd58accd0a118948:a0ad33b985713d8427a305ed3422286e:5775a82ce76c	M	t	7B
810	68434842c49cabbf563fb912:cf90399d388af5af11462ee2399edd7f:2e5a50fe	f73d1a12d6463343e64adda5:cce423968f61efa71ca8583f4495141d:6cb2eb1d49b64a71	M	t	7B
2042	86efec76a59820a2d0b0111f:f7b36c41d93658874690019235396e78:ad6b62821a3ef70edf0d	261f9e60e8350202950a5660:8b814b26b55bdcee3663ca27bca5523e:8356b92de7	F	t	7B
812	e5670c3c65022f0de63631cc:9b3d10b59628fbbbf8102a552da28c78:a7c40432	0de512ef3d1835f587a75e5b:dc226f1908a808a77ed16b77a61241e3:640f172b83b8082740	M	t	7B
1800	6c573a4919568daf2283bdbf:1ca8204b7e873e0d7b1d097bc0fd1708:0cc5b302c0	0b7392584702a38fc2c595ce:3cbe43b20c472a6e19a8311f096caccd:03c7e47c33d1697292	F	t	7B
2296	363ebb73c0b0c8f2ba0934ba:a25f44aa57e2965bf8b7cfb2c8aa376c:eda46903	e745708aba7eedf0a26f9ed5:a321f192d532bf6ebe37e1ebbb307794:19bb7cc0d2	M	t	7B
1526	0c01e5008faf5d64985dbcc1:98903ff8fe947593fe53d3b93b13664f:ecbfcdd97e	22e4190f2fbf67f23270a2fc:666f8859c50451d6db639d02b936fc2a:33b2951f5095	F	t	7B
2561	d7dfa3caa2cc399305c61f6d:f4990c374e29e162abae6d242d10b85a:88b0c5edc82b	ed4bbc291ceb900b33dc9f6b:48465c6d2fdc9f42a48d4d18bf412a10:05ebdbae65cf	M	t	7B
1554	fb75634d04414b5bd1ff3bcb:8ad07d60eb1dc6e216481c4ad3e6b1c8:ae642d88	898963e412881b4feb160474:f3f1d8a6f72337454b93642778b41ef9:d6a896cc1b2ea6	M	t	7B
2592	8f745865eb18cb80c7cc2f72:3314419ed1d1f6b011dac707b63a662f:eb3cac6fe76e	c506c0f7a9122af54d07755e:5a9f014dcaa81e69bfa7fbe595fe39f3:365f4d012ade10c508	F	t	7B
2422	f75866517b91dc1957fb9f3f:dcf6a23f515c84fb0705254a86c4a006:a866493ab1d2	1f3037f53a799fc320f7fb15:dc625435b50a0e590d8a474462817ebe:9c02a4bd686972	F	t	7B
2562	6a70b8af8868e193e31834b5:45f350b6121dcf7c4bb4de24490fde1d:b604582e	d7afee6082b4dd11a4b0f5c9:ebac2b8b8c94b68119ad1a8b0779e937:bc97c41268c7	M	t	7B
819	08e39d110132398bd6c73c14:6bf1aaf4e7eb89521346cf4805fb28d1:009de8ab105b27	ace3f2448643ae1585e6b410:a59f1e9e032da2104fa6f3a4774a14bb:a53d53f92b2e72730669aa3b5c	M	t	7B
1244	ec9f362afd5b3b2ebdfa4741:e9fbf77312a3b453b8d87ea8083de4d6:0c90097a03	c4717ca8255fbfdd44e99402:2e5ebed3a67b562553181b5df2138f02:142ace486919	F	t	7B
2207	d90dbcc22676f2625be4fbea:be3c8a0f1a6e70a102d4f9a8630f312b:9eceeea446	d0a72d7d3afeae0ee8e854fc:a2f1703f82483b68f57eb08dcbc4422e:9b46f099c6	M	t	7B
2356	2bb174ffc3825c705c8a7029:1bdc776d94fb016f4a94bfe7bc0a4903:4c0463	f692c0cdc8bf7f791f8c0a60:5a4334e1459fdd8517fd6c3db9ee5b22:4ccca308f2	M	t	7B
1057	ef330937065c463420857fd8:310d47f5f1f5e6acdaf9d14c88638f8c:c587a2bb89	2bcb719e344c9e3514941cdf:b999b92adb2f6cf84fac30af0255657f:e6328c6910	F	t	7B
1689	6b4bba74b4dccfd53b3cc821:a50b09518ef0676bd6223efd51433407:443805b7	af940b988c59c2d6e5e2f953:13d063785fd32f97f6cb4ccd59d7a6aa:39bdb341a03509eccb2d	F	t	7C
2164	Ihsan Khadar	Mohamed	F	t	11B
1916	Ahmed	Niazi	M	t	11B
1917	Carl James Ardales	Parone	M	t	11B
974	Cole Robert	Peterson	M	t	11B
1918	Sereen	Saadeh	F	t	11B
1690	5f0f0103036ec23dba12915d:8334dc5af43d7ac38a7731942450af4b:16cff3613a	0358f264c391c634be0768c4:38c72c368cd5e4cef3a70cee35b106b8:a7768938ab3d5dfa51	F	t	7C
1053	e4115d10af2661a58250b240:c67ba0471d86a47e26311ce5bd068274:fba4369b	70542aa03111ac2dafe43b1a:778551057f03c10443994a6f91f77d5f:f304bc7cc387	M	t	7C
2405	7c61c03f87f1194674155f1e:7dd169951026ecef90883f6a5f01290b:0e615d794c	24e8681bca9ae08e32c51e1c:00c385c4e4f342d5c01aaa449b713c82:aa7b1f60a04f	M	t	7C
2257	0e2e38a4ad0d96bd494ed0b4:ade5b2ffdc814d5ce7f60dd6bea549db:08d63904ff	ea7418c254fed5fc1afba18d:138a1cf3f6341d41d955ae952c86fa7f:131741465f6c	F	t	7C
2312	e35da9910b9d588daa1cfae4:5ec43289648b5b100c19acfba7f4bc90:66353e32	2377f5c21e9363fd361d252b:c7519a567449875f642966c1310dbb31:f606cd89c48ecf	F	t	7C
2411	0ae66773eb90a41cc5260099:2c41bb6e961ba559a0e7af7493d792eb:2526a9aa4a	ba19534ef4d9a61a0f7f35ad:db08f5988378ca4fad87bd5031128fce:fb6cc573abf516	F	t	7C
1071	b764a1a8521aac2695e8c54d:f7f8e0ede7d806491e7db40b0fea373b:a6cfadf3	b052e75b4ef0fb60f0887450:0b7feec8407cdb8734a0dcf83aed2297:057dc76be5	M	t	7C
1955	431975d55339a075fc38138d:38d2b650792c8d2c08033ee7f944d87c:d18008bbf15743bd	206b0c9bb4996fd91fbf85d9:35d181adad3c971f3eb465c22e0313cc:79acb0a887	M	t	7C
1804	14fc1ede5bcbabbeaf2a896a:ed619034a603fd29c5a754ca234be35c:ecab13cc0c	3f96be2ab9bbe224d3a4c871:597ac32a25c7661ac81f7e1c2384a632:ca8e04e7cc00b8	M	t	7C
820	cf2ae751d154ba9d6f2aa744:95a56c662cf2678889f151c5007b5b74:a6308213	2ceb1b55abd7d3bffc834bd6:cc3bf4ee44730ce2b9caa3e97936f330:4e49e3	M	t	7C
2334	566cf63363c92933aa6e7c44:a8d003a71d711b1bc0aae0f173e880f6:67297ed71cb3	203cd811dd53ffb456b00cd3:40defbf79f8138fd90265a659844b8ab:c4cbca5fdb3b77	M	t	7C
1389	0689e3b4f5f1f22dc2ed53f7:bacd3dd6fc500f2180bef37ac432ce28:c4c981108c1bcf	8e3030d25d29e701d2c05248:e19a6f4313ae4978e8832051ea7a7631:934007ad5ff6	M	t	7C
1571	e9a0a4306bbdd11a67eec8d6:c07de5871aaa2986f3326213f9083d7d:c50f249992	fa609b0dced007f91e61e41b:8d80874876b9850a1977443e310bbc5b:47d97088f4	M	t	7C
1117	c7ffa0022c6376ca79dbe262:f3559e2c514d2553c6b5e4de4e7df624:a5d8f7c9ff54356aec17c5	343a00e7021276bcdf11421d:90fd30df546cf57e7fe53d304b000e8b:5534d17d2e	M	t	7C
2566	7102ce10b2525c5cd95c4f8f:a77dea395bb49a5a8639c99405c8d4ba:4b49a62b3df720af28f7798109	aa39adea6e0f036987cf651b:44c7b6667c197986deefdb2859bbbf72:51f814929f8a	M	t	7C
1594	7ba198764b007c981c3676cb:b5162d46504518fe8a1ce9400689f2ed:773401aa48bd	15d645a5a4bb641b09e7aa9e:9b7ed645e03601db8022bdf0bcf317c8:989b8fc8a10c	F	t	7C
1593	df87d8f2b0aa686a99848bb2:8b00ca2b90153dbdc80d23c5e1e203dd:811eb9391886	9451c6ac7177a6d0ee1e7a43:164f9f4001142284744705aa671e7382:15850a69bf	F	t	7C
2199	4227f423a6fd34671b8e7073:e22fc94af28ba7556c7f3a6eec71e2ac:074a03dd250f22f3	311c48d490e474eb62aaac62:55a41864b6ac9e7459919c0a053fdb9a:4e1c15f91bb2f9	M	t	7C
826	883d312ba7a8242a3b763b99:8210215148d656778c6aed6cccc8112c:f16e3f80b015	0ebb0ac444864fd70ce116fb:ecd48b700877ffc171b31fce2f0717ef:00f6dd	F	t	7C
1688	f965707b68f3b9ed4842788e:cc088aa36bf5d3e0d82dd3a953272a8f:c3c7461ded	4a0f87739569e280379810e1:892ac27f0f342ffbdf5c9db87c8ded71:0da01c13187443302d3b	F	t	7D
2131	1853ed78c8290e9d6687d951:f1a6ddaa6c18ae29731e90ea6d63a7eb:86be109abc	38315ff9a4370b98913e4e96:6bbb9aa7b85d339a44764c56f64b8926:32c3ae4547	F	t	7D
926	6f38cc507f3090faa4c2a0a1:059fb14211366f65b7f45e017e96525e:a00cdfc0	1d4c111a853a403677660c0b:a30b731185ab5622a66cb40f00e09ef7:44e4a03bb194bcd8	M	t	7D
2291	2237fd5bbff6e38be5725348:85f05964ded57e3b7ce583e12e5a6489:38a293b18684	9a0674c202875faf15cb1422:219e2f42d61f3379e16888b723354465:7f5a71180d	M	t	7D
831	fe127dff7c92d34310649bee:f1f83cb42da95034614d9a95c3df8068:54e03bfcc6c0	b95d1b08234cb93b94b56dba:2a12182fd0eadba84a0a97663141e50d:b5d948df4b18469f2bb59351	F	t	7D
827	44510e9a124bc7d41debad29:bce43b8273f3cdeb9df430d493db39c5:744660bef6	3e5ced583efffff54a3357a6:4e5225b31747b2a93b5cc23ebcad2d81:f6301036e51f83370521	F	t	7D
1246	cb4b3545e1e4ee0230578021:8b5bc711c5c2ef9409ffd9481da0d3fe:6ccf4403eb	9b51100c2767869e22e1d5ce:36a15992e9d2c171860cef0dd6d4362f:c6811a8a2daf	F	t	7D
1247	ae93c85db1dbf2c0d7b5c568:9def751633f4f6395e21d112292d715b:8c711f129d	58c1ada9fb49431eb0c9eb09:4c5045796f0208201541489081497196:3fe74eb742	M	t	7D
1531	bad5ec4e67749ec06aaaa2a7:e278c8555eabe2f4fd4d7302f13ba0ca:10728de1	9f24998f72b3d7969a64ab7e:74190365881b1d90d59fed93ce236784:6abed952d9b70ea36f	M	t	7D
2033	3d49425c943a1506989c7963:25feee88a35ed96ec695fc593c190850:9eb581473c	b26fc1aa3ec114311d31c512:1dc246be8a3d66fee112b9ec47b0cc4c:0ddd0b3f827dfc	M	t	7D
2252	e752249ee811ea7b9f463872:2ad5ea3df2767debed1d700cc654236e:a89b3b9b73bd	8b80b9fac3d7434fe11c6444:e9013321ffcdd2a519ccad43c972a0c7:fbfcbc51	F	t	7D
2563	5c70eb9814c0dd0fa7456fc3:51cddd0141658e3ded705e8b44d5d00b:744eecf12cc454a26b24ff	f709c1f1d8f1162ad99ecc0c:cacbc0cb03067455173f3c6b2d2a4c1c:6ce2857bbe3d42cc2b	F	t	7D
1897	5430ee5801591e8f5f61b1bb:1fa122c5913933177cdbdfb769cd7adf:e2ae8e9e	d4971cc0045c6d1eaaf15fff:b351092074dc9bdb342420995a8ff009:6e0bd277d6	M	t	7D
2133	90e7bea1d12e661e59941736:8bf08c8e717323eeb4505ef368957e68:21ecf5af5c2cd7	e4b269e47850b5885c28a87c:dc75c62bdcdb21ded1ed852119b7aa31:c6b6e090c4	F	t	7D
1118	7e8cac99cdecf8ea0208bfab:c477e07b6dcb50e86fd621bf37624813:2a2e57050f	4454c54188b0098b72df6b6d:0c818ddcef737fdd1ec24e7a6054d847:ddbfe75c2f	M	t	7D
2565	f6303123d589c57c0c8eb349:8419023dae341b60e692e5367ed804a9:2d45caaaf6acc047	17d71aa909c1204abfaaabc6:176df42e8fc678a479dbe4d51725e10e:f62fc502d4eeef735ee1	M	t	7D
1367	9dca3b4e4c7b0c6974f0d1d1:a971932c591f5ca47aa5674fe52edcbd:fb0376a3	ea9db343b97c03cf572b4f18:a3db2e5c4f2760ace6afa10612860dbd:f76f0cb78b7d	F	t	7D
1073	c2c5e58d262764a3658a25ed:0a306af53087775ffbe90f4adfd89cca:ff6a076223eb8029	7729dc6a4e31b19eaba7718a:a2b580ce17496fa9f4d1cc568da80582:f897b6c228	M	t	7D
2377	da79e841a6417360fb0e2fd2:da3da17961b83f6c1e7b49b8084d7702:1c07a88e7584	417c75efc123080c803f1e93:8a7af61c0769b70eb9a3273c91a7c8cf:7b038576846bd0e4604828a0ec	M	t	7D
2285	ba9418facc2d4f7801137392:6855ce2f8a215988ce8a3ffd9e95f537:4af42d7b	e880a4891127644e71dbb8f2:f1bf083f4e1c0ec76ebbc5aa5b496d71:e4687f98bf8297	M	t	8A
1763	1f887c5d783799a7341dafaa:b677d34c94b17a9060b12259ebc46bd3:452e17145ec1	dd0bf05c25d25ad7e3939ba2:8dcf34788f4730b7f6f1530d30d7fe47:2336cdfb629ef011	M	t	8A
14	c83e026e23853202f4e63060:e8d6a5a788961fe50f0d6a7a4551f8d6:93a93a9d02b0	84c34803a685d8a5e8b65f79:fecade199da5064da6061b119ee596a5:83769f2a971f	F	t	8A
719	e03695bc475e5045e2d7e1a1:c05053fc6b2e941135cac734e71affd4:92576954	4e0efd3198cf6c582e191507:591dd2f3c058bb3eac0291cb2e319321:d7a7729a56a146aa152686b4	M	t	8A
2567	4bf8f2ef283dab2dee73fb3e:0e97a2b3023e896a4100477174e83956:21be94dcd98e29	a7394a6d8d91cf8f8380c9ee:f72f646daf2b69eff63bc580e245f3bf:122635b24c8bc9e249	M	t	8A
2568	57cd7e0fcf9eddd663b32c28:57c01aba3ab45fc3c75b61ec2a4b6ccb:8e55aebb28	4414ea63d7076a7299e65593:159dee1a856e7cc4b9a47abdd91f0a2d:28c0faaa5d6f0f	F	t	8A
2031	3eec223d0d11720183c8d975:4a086cc160b69bf5c86db90ebe1c1db6:575f683e9e87ea	555050b85140cb5cd468b023:c07f46ada807a5f6900140674ab39e8c:9fc214bb4a96686091b9ff3b5c	M	t	8A
1997	4a6ec3bf76eb7fab7f58155b:4895e03b7595e652166c7ab81d1d240b:9fd3407b87cf	6ed14c831376ba05960c1094:a3797f5f9163bac97c8fa52fa1e9f49a:d2014f2b4e7a725895	F	t	8A
1812	5462d7f4c9ce83de61004b6c:c81c76747b49723ee377404f8bff489a:28861ee47a	4889788de9947a5631c1b01c:6aa3228fc740ab3470cac54067c651de:1969c4fc13	F	t	8A
1901	9bad77f7fcd17c2892510839:564ad3cdf82c0966cee0625d56724564:abf6abb70d	e9b7feb3d35d1f22457cd8f9:dd4ff199feea22e3988ef1fd0304193f:b886e20296	M	t	8A
1816	5b328f7dca7fe87808200927:3cab68ea6b7b0930fbe09a1b80122580:d2da7547	cffd3ec543c3e35818399e51:2400657d4bda75e00285669eed5fdef9:f29c90789004ec	M	t	8A
1701	680aad300fb0bd42e17472d9:18e21b25ff75fcc465bfb965b448b8f1:6cd9dae84a52a1ec5a1d2dff1718489fd92d7f39	288c7917a8cd1be1186429ed:495961b37d9e0c72e771293288fbe47e:f6ce00c24041cae496de4150d2	F	t	8A
2184	e51b75562ff1ce6e74c27a60:2703ce4dc8f0ed056444675f4dfbf214:1041b333b7	513c5fcdf0f9019828378c66:1ff3806a736a9d6ce8fca5108a54ee4d:8ce0498729b450859d4c7e	F	t	8A
1818	58687365fba3264a20082a1b:816ff00531067af07ba25b9e29d4c3fc:fcb985e3cb94a6e6	a84cc931b0294fe401ad43bb:99c973d71c6482794e14c5bfb5dcf688:3508c0c7d5b0af26	M	t	8A
42	7d2cbb434db2978ceb401e69:43ea814086ba6a699964e3b71392e95b:dda0bb1c	414e19170ae25158c6986c86:2e76882808d8895942f0a000505f3e87:b5433ed0307777a0be	M	t	8A
2007	d9e328052aa1f7ee96d6ef29:c52c7346fd8ed8eb4ce29834ac9fb8cd:ace7ebc4bec5be	458a8b3a07b6722258e631e9:793e1053698e60ad3dfb7d20ea038196:704477d1ea22	F	t	8A
2038	49cfe50a0b4e14e0559c7500:33ab7505936db849c3a6ef50ef4a2065:23ed7c96	115ffd3ab685bf7dde716b38:ce7a2d07b44a456f9c9be1fb13c2cae2:3b7f41	M	t	8A
864	07190ad252910e6d13d1e054:a01d3065858d97f8880e782f7247872c:edb3ba3c8c	54e6d59b58c521e1c0774e9f:04646baf93f194a62b7b0f368918921c:f1c304ab84789d	M	t	8A
1697	0a45a8dd751e7ab143c575bf:263081d3fe272b6197cdb32edddb48bd:5246b39e65e8	dce1b5928d923667a0dd2b86:e448ac45bd13d1a027a5711e47cc5cae:04a9d26acc4c7265fc	M	t	8B
1082	cc6c5ef48d4bd6aaaf22e952:c1e73e256acf996e72fffc3f550f2b81:68dd4f58	3e23cc41b44589a71153d1bc:e8d2b7571d801a61089932aa87ca397d:c1b6328520	M	t	8B
2186	e8ed1e2d4e6b727d89fd16a8:97d6c30f3fd31c20f0b8e3a2c7f67ebf:b5d2c5e0	e0c49f376dac6732ae1c6ac7:c17e2406316d0adeb9dda6f159b0f315:f2a7461d459f53ca8e	F	t	8B
2362	7b19afce8a2a68c346786621:1bc36d359ddc5c1a79d52bd2313d003b:6820bedb1eab4f08013d4de4655b67fc5d715058c681	65c0a54abfa253e1525ada67:c5a252f0d7620d5cef6aa1db5b058d85:1737ffd2cef83e7d42	F	t	8B
2569	3ac77e87263264afb3e6b976:47cc255b9bb2b986927c4b076833ab20:0a0daa	8a4dbe8e200099d963fd5d59:a5796c09872f1374f47f7653f5f22573:7ef2c8811152	F	t	8B
2236	e0119c8f0ad741c20f07e719:926988e7088000e2be8ef9ef8fd1c561:154d06b48045bccb	29f049d03435b6dc801aecf9:3a35b19c2847efb3c5ae662121284b82:c5d0ccce068b	M	t	8B
2040	26873cdbbafb314f535f4be5:33ccb0bf584b0aa6cd64e1aa8427a2c2:bb0f974a5c4928	8f4bdfafb1fe95921b9a55a5:6d915e82a35c50119951295a5efe661e:0fda2a320eb2f7	M	t	8B
1250	5d2f6d8e2a9c758d75a2ec84:17053259aafc98b77b17ceba93597553:1da3656a9b95c6	924d204e8a8f29adf8838d58:92491160081c082660f419d6f06d40d4:c3e2df7d703913	M	t	8B
994	7133aa21ff5e1dcce418a5cd:f67600bf44b8368d2fc86d9f819321ca:a40e0100af	d57262bea4b7ff30ce2f4288:b9b763b656e33c8ed2903340cf7834b8:e4213942cfe6b2	F	t	8B
2632	5d8a7f3760598856be26b55a:97377f5f4fe9b17fe20e4de1954c9ddc:e281d599c5	471efbfb0be753ef0cd38ad8:b066119a9593d86fecba0d04c86d9dcd:faf0fe74	F	t	8B
2665	42650e20d4a0a682a2bbbc8c:9b86000513aae85c206abf783bc1636e:8e89ff66f8259cdb	3368772ffc682f087ff71283:403ba65de49a164b3230d0e9fa43e6ff:f008d655	M	t	8B
2615	9500588330a2409b6665c341:c51e0576e1e661f6abbc140d6ff54437:e55c9e8819a224	95f49dff3c6e478437ce6d1a:4bcb2e72710ddb9db4600e6bc855aacf:e1997518dcb1	F	t	8B
866	69ec9cd526f5c7e887696071:9687177b5cf1d60c97a598dc70e32330:46d731bc	a62c80303e4aeac3ed179689:aea10fda6fb952d157666ac7fb689aa4:4dfe5826dd751a	M	t	8B
2140	207329f2468268c4adcfd293:552461d36de040335bbce72100a58049:4cd83f	5a1dd07b9882b67a46eeca85:2f0904cbbd46cef7fa107c4b4c625154:3324	M	t	8B
2145	b3a87c757263613286c62dc4:2fcfea8800cbd001f3d7f09d21daef79:c7593b0c5fa8c8c2a0c126088eb5	e2b7847980484ebea66bcaf3:1aff5e82449b33ccc15af6be10d246b7:ee0e5722a7	M	t	8B
1904	82b10d75d4b85544b50b7240:c5bd7a28025eac1be362561fee58c7f1:0613779d1181fc	e386c9bb369af1bb54ba10a3:bdd00cfc688994c8053351c6c9294b02:4714a368d6ea6b11d9ffea	F	t	8B
973	981285de1b353c3077c4e083:e1f9f50833fe716c6542056a03462326:fbb00b3695c8	5d9eb3854e0b0351fd7a8532:ccf7c8d7eb1a4a9187132c4c326873c2:948c3890490770ec	M	t	8B
1140	e7da326f01a576a6a93629ac:d64cf868371a3f2c8ed31f534342c2bd:3d766d3b493a	054398942fd150241cea8774:c8540552e032ce247c81b78ae8589720:99367e1c8ae143	F	t	8B
912	e615803d5b3acf335a033a69:e32890ff2c6c3843229f30bfb7d4f03b:a9b2ae7fa8164cda	c95f2c6eba7eebaaf7ec823b:7adc1a9a175f0a38027a295645181619:ad52518e4c43b1a4	F	t	8B
2204	8f0111819d0c824ed4b3aadd:824423f624dee7e5bb137429731a598b:6596ccd8c87b4e	a25f6c27e97536107350efb2:9b69fec38dd7cc7f6bb47ca99e8fca89:6993e2d99d	F	t	8C
2586	818cada1608afb7c6dd6006f:610542c0202ae9688ee4fc87076a9db3:fc528d78	e9c2efa1598f610a1a902d65:e8b10472d9f5eaf8be06661f6ca59411:7c226660bb628f02	M	t	8C
15	d7dd9cba0e47cbe5fa113b56:802300c6fb0ac3a6797eab7d4a5d9fe8:eeb2c66d	f7afa5d139e2fc4091a22086:649d4a62ff24e524e8fd485169e9f137:1a166108b88f	M	t	8C
22	3597d2414e71d7be7fd254fb:0c1796dda29026c49e9845e361032aed:8784e26074	ae897664022368236205729a:68f8e5e71e416ed87bfe15410d872921:7446f4b122ed660e	M	t	8C
2570	fb6b8d220946ba8e5d729f56:2090740be72beeb637b91b417cc84ab3:8025e647118d	aa92c7dfc951b751a0db3f84:26d337452ca591a42faefe471a48261f:65419581	M	t	8C
2315	2fdff4fdbc226df5ffea424e:f741bc3f44fd76d644302674d158c6e5:dc3fa22d	eb436584877d055be907c8b5:2879cf9d3b505bcb2a3699214371fd96:3864a89f44	M	t	8C
34	f0ad77116b43df2be1810fe6:12d1ceb22df13e7bf9091f1d15aa9194:89300f2fcb	881d2562d16c549cacc95050:b6bec3a96110706a07ed1cc31bd6a35c:80b61b8345fe	F	t	8C
2139	9fb7024d0860d17e0ffaa97b:e0a5f9b0480875a983ed0712466d51f1:fa248c5bc61d9b8a	058e6d20cd1dedddb3a2d200:ef6ed21e9d39450ff307bf24a05eb2e8:bcdc2e2d8f3f	M	t	8C
2374	9e0031660b9502c7cd9fcfdb:95503d24c135d8a9400eb30707cc803a:eeb1245304ba2983	f0054e64ef27f4313f2d9e1e:0f70c7e00f165222eec8de6e0703acdb:892a1128	M	t	8C
2364	bfab330217f7148fc192a9ad:a2256e1ad717168228d864461fe7a75a:b0bf37dccc	e0251a4693abfb75c61e447c:60edaef060032983e413a3123147d127:2829	F	t	8C
2658	9bc540a16a6eaa547837c455:fbc5c657bad95af3cede19729d4d3816:1df260b6187a	27b537f3fdc8f0495ca57f33:ddb438c9714a1b00e9817a63a2403a47:7145	F	t	8C
1259	113380ca66f97e638fecf0d6:fc24eef0fd8d9229cb39418a4c01727b:b0a8ac70345b28f4	36504995c1880a1afc207d29:b0d0a1853e5bb3378751ade8e4bba69a:338d15864dddfa	M	t	8C
732	Abdullah Ahtisham	Ulhaq	M	t	11B
2588	4ec0eb18813de27b7263dc36:db198b41b8d512d3a38c73ad7d4c106e:1358eccd3bbba2	68ecbe42c012f710e456cff5:4c07e5b05c3132d4ac2c629e3e35f61c:21cf39b52df168fb43f8	F	t	8C
2344	8b67dbb08044726b7d1b74ff:5180ed14e5a5d164399be4c79537d36b:e2510ad2dcd8a6d4	803b72e180c0314d66c091b4:c70bc83630c2aa32f354b549b1670b69:671f169352f8fd	F	t	8C
41	bd0cb48d59a3ca467d32d837:6e391b9e504dd3622cb908557048b6f6:fdd2c73fe3021527	8b49be0e03077e411d8d9dd3:bd44900dc8f05130e30b7fe159463115:cc6a475fe78b9d	F	t	8C
2367	6800b219a8ae27f3b57d97af:2a7212e965a501b8f3fe6bb94ca1ac47:c9b212e3420687af	c2274f7472478c4b8a475c2e:0a54b2eac2979d63ba5f93fcbd738978:35658e1b1c	F	t	8C
2399	d2e02df36f29ea2144c63c9b:95fe3e0ae187e2caf6f961f0ef69b73a:a250365e27688480	6b57497a1c749c85b2eeb4fa:1a72559ac6f2591f703e8b0100867d72:964e11db	M	t	8C
12	bb4e56a128208a448b38c1d6:cc6015f2a05ad9da648da14b5a90c11e:48fb5cb5	8b84d9b385741a41f0816f40:14af8116bff3b13544b6d5365a9a467c:1129eed4	F	t	8D
861	88e23831a65c05eb2d7821ce:6e80f47c4504f5c789619734dc541299:adf045d535	a5d82e190c05414863e3fb21:2d7f2ed73b0a739cbf5db2baeb0c75b1:000e8bf4	M	t	8D
1457	b7539a02eb874b8e407c0e03:69b4e57d5ca830c2bdfd7bbe59bdc90f:e843a1ba3e44e5	c02a890981e17391bc44437f:2acc9acd39cc56bad965f72799175094:7563bb05	M	t	8D
1599	43d287825fef8d6e31e59f9b:e3ff1b1496f369f5059ae8cdb6b7cb98:3a72d5	e4a75f670de6ac7a254a50cc:09d4de0b968a03aeecb1513c3da5c72e:05c653475c5bb7	F	t	8D
1769	660a725ad357a67761280bdb:9eb15ad185e258d84ae5f6aff2b1ab42:fef8bf10	ff7efd652ae427e0917a7c8a:1cc5de9ad69c0f3439f268d773125b91:92db6a88fa	M	t	8D
2309	ff725e8cdaafacf26ef67eb0:641efc161613dd8ed84c3c34dc53a575:3952a318ed60f7	9baf629749d0f41c4ac64297:9e67afdd42d0e0bf9d567b1cb253f754:8adfd6a84ecfe335db68e3	M	t	8D
2571	5f02ce884bb425e4ef269a40:b5e5c06bc3db3e281ed72c272d5dc0c9:18c79a2b	361dbae257ad525570bf6a9a:2f5495772118a4749f5437fe4fe0df9c:ad952c900a2f77	F	t	8D
2316	05c477408afe5e59a7a1d270:c0bb62617bc72fea0b873152599bcf32:2f3c66e95607	47aa4abd34f8d6dfb7ded990:933efdb515adf114c4d002d8281251b5:46f6f5ca8c	F	t	8D
2654	52909091507c1076e1f6c6a5:14d8ccf066585d7c3cfcb300a75ba0ca:479ff2094eef1e59	b45a5143626a10009ba9102f:f0fe6c01fb0a1d2f5298d4ecdaa2a8ef:a3b10483f636fe	F	t	8D
30	cadf270b934058088ba1e297:ce22be0e076fb58acff0d4f619036a8e:c47ee98f0be200	74ba57cc8d003cd2ed3e3f51:6ebfad223f6ffb10af74db6f4638b992:01a5abfffa	F	t	8D
32	0f99944638ea62c09bf62d97:76fcb09c1bfbc1638bcb4d8d74b817d6:2ba961a45f1740d24d6d	9a440943add7ca163cc52f2c:fae161bfdfc43130efbd1b7567cbb3b7:b5d11dcf94	M	t	8D
934	09c8e2444a24c509baba48e2:a5c48da7dc52d9ee59273519de6a4df3:b3353ceb90	d79c7c7fc9ef785ced10a249:fa2a48d6ea9dffbcc1c71ae2abd1cd27:16d29078	M	t	8D
37	a12b567a31343b77ffdf8f1f:0c013380413fccbcb6d2796d79cc2076:91cae8911b	3bc2d0a3e8b06d116d621d60:bfa8a7bdcb7cd3770684516fe4b7b7c5:d2f733b5b6e58c	F	t	8D
1702	1ab4cac35a1411b448292cd9:6f54595c7948e03ed0cbc4f1921dcabc:ec9d8456	694417ea3680e0cd66f91cbc:12a8b2ff5f3439f1a8760d8452c03141:5f0f5d35e2f2	F	t	8D
2141	f7c5f026cab1b8169f1e43af:3d637fa1f954bf1fbfb4781dd8b040cc:7224b0f7693cf6	6cd650dad6c8b364771d8142:c8fcf39c969fd63de554e8a68b9858df:c941f438	F	t	8D
1902	ed5bac4181c5672c02fc433b:e70dac8f1797af4a9b3aedbe3d8081c5:c8e965fa65afd1ecf604967c8d	1b532eacdf1229f8d76114fd:c4f52453f1e90704121ea565fae7f5bc:72074138de67	M	t	8D
1703	81af395fc338167909ffdbd6:3068ee2095eea1237772d9fc44289101:1cef08d196	09b560db3ec87f58ed819657:15a09374de66329093c3073b2337b4ca:11ee76cc92	F	t	8D
2006	98f2555d61c288a22a32ac28:281eae0d231750bf4958fe6dc3251097:210cd17b	232bf8ff90eef284062bdc45:bcdf2cddd5837e25de70ee17e8e17f02:8960e3fd57	M	t	8D
2572	0d8f424227fad574b9b4d967:137346de48ac4b9edab3d6d3a2abc509:66392e0587a7fe25	ea591658d18cd42987f61ab1:468950d6c6169972e27090f16da91778:c4fea0a07ae89fee91	M	t	9A
2406	756181e7ddb536c841d1f513:a5909b9b8a0d88f71b95c8dc99a8f01e:22c3db5da6bd26	14950f6ddac626641498fc17:933bdf7a37e55b2254cee6ccbe587bbf:2b90dfe81c4eac5398	M	t	9A
1405	67f990400507921cfb897e2e:e5d6e0404f5a1dd2bcc9de5ecd7c89b3:0c4b88fa6f	4e9283281074d0374cce70ce:dd2d3a2664b5c8a2bf6dd10bac6290f4:be8f7a061ae1db	F	t	9A
68	a3e8e326b1dd03a253d36b59:09220174b3e224ac4ea3415d7d9a78e3:7e2e985c	0c086b673ca4dca42e0e7001:c48562b34ea97bbff96bf516fa0efa13:49c1902835e4dd	F	t	9A
1187	46078c095c361f6cfd3d7a44:68083dcd8f28182b40ca679f25b08cf6:eab8f982709059	d415fe41bdefbb9e04276da3:40816087be016c3408afb9b26127b0f0:e964a1b07f963a26033b36	M	t	9A
1943	ba0e5cca229df1519aa9ac16:801c351e749513c16a9d731d7d768b01:9d4bd33473fe11	46232dd5ae1d0fbcabb618bc:a313a7e0f775c361e9c4c15496590904:58e87ae8	F	t	9A
81	84cdf36b7b09d21538aaf210:ed65ed1669c7148953ddff3f5b3725e9:3de98a5f	b16f01e09b5a3ea8f838e1fa:48c726f878b2976854081543fc269c2e:31990e70b2c55d7e	F	t	9A
1540	8a9340216ec696eed17748d8:88028127a1337c185551415c21f16fed:3f6036213f	ceb6fafd8bbc1325950e17e9:3640c72d6023d33912239fea43f7ebc4:a06484226240df4e91	M	t	9A
1538	be17697ebb54b618af9c5477:cb8978d9c463c4beaf2190586859d22d:3e6fddf7e707b5d9	5545fcfe8b13e41411abeee4:a40c4b9e5b82ec632d33e8d3af279a47:955c36e6fc	M	t	9A
2378	0160c406717159a72c01aef0:656582a3bf47364dead5c6c073de4c8b:15304f234445	f54b024bf2255ea520f8a132:384fc14eb2eac55bef401e4dab63246d:e6baa3ffc4326221	M	t	9A
2320	7a184a0f7b82705754a40668:80e3351c876bc0dc5a036b8090bf033d:4c9309bc0c	68fd6b2dde76f2e0025be3e3:28690d2f48d1dddc18d173d85dd63ed6:d564f46d0bd3	M	t	9A
2240	ac3e1cb78b17374bf38fdec7:1aa5634ff1d87e4eb5fca3afc696735c:133ab45c3df4	7365db74688a867892c211df:01abf4200602e3a446a7404a3952800d:000e7cc18ce022	M	t	9A
1764	c5fde52d7ff7fc33af7c7fa4:f1b85394264ab0f5ae377b4e54abeccb:07e968	a11713f34381b9a5424e68f4:15ce0c68058b389ebeaee1e2ecf99fad:d04aa0cde59da4	M	t	9A
2243	79c69171355342ca7eab5314:dc19c362e604a1b72e7f9010e4a4b6a8:0350f6	36c6d8705b15fb96fff09457:8d0a36eed0bdb79da0180d4af54b395f:33a720c2	F	t	9A
2576	5a89abd450ae0a2fe180cf6c:100ce5fe912145ec9024dbe6a24d360e:981e4b117b	0f3979df264f4e1540fe5f89:d6b6de16f1210b7dc2461efb21254725:739995c62576ab	F	t	9A
100	b976822101bdd2256ba65707:b9bc3d120f7a73e37e019c14b1aa7e4c:c041891956	1ba37f87871458615eb718bc:4ddf4496d985987bead205bb2e12537e:05a555390bb6	F	t	9A
2577	ee8af271b95a6430d5dd013d:68ab6918684b311b5034ca20081f9b0b:69e9e97cb6	f8edefeeedb33296431f5b59:b0653533dc07e7d18d28610e49fcfac1:934ae9548a75e5b9dc48	F	t	9A
1191	32d2f8736da99b4725571639:d16f061fd857b0f565bdcd2fd741f4a3:6ce873894a3c	8fc88d0a3f416e7c9fe2101f:31c63a2e7a4c5973f285cd12c42e5652:f8f40084e6	F	t	9A
1198	570765a89c3a515d9352fb01:1e397d820cd8fc5f99b4c6453281d2a9:bc43e6d5ba769881	4d0870ae3cd08067b8647659:10f6d0189c93b952c1177c096a56af16:b167380de003c0	M	t	9A
2401	9a1dd2800163c8b026061204:5addfaac31499d041ae3194da5ddfb09:bab831c8c2	f31352aee8df9993bc6819b3:4d16ee14d1bf8c241142595790686073:35cf7e2340	F	t	9A
2147	9aeb82d988e20918852a54be:236bae325408a0dd13dac323268066c6:d893ff322110	cb58f2c59b2534a300cdda9b:fd7a3f400276da2ff22f5a3c9efb5754:428615275b	F	t	9A
1194	b259fb55a4363bc97a324a52:074e48f89c1e22072db1be9ceecf89ad:795a44c14c	be94fdd40a3a3dda7bb9b97b:5871d5b4c739a417e9a30910bd88c6aa:8fb049651cae3ab69d	F	t	9A
1129	e42c9392c673b2aafb765887:deb9de5ddb73e7e919851279273fc026:5671d861d9	bcb5b9a43a8b28d9f1beb09a:06026747f235675036de7dfb0a80e62e:2fffc1671eab9a7dd8	M	t	9B
2573	420a5ab9ca03af10b75c5f44:23f701e2f6c99fe2378fc4a933267170:fa0613858b3d396d07	1712fd8698f78364b32cf292:b8d124cc6989d74fa90e52f4de3ae1ef:bbabc5	F	t	9B
1088	ab039c5f0993f7478a868b9a:b41f4ba0e111cd1fae3bf4311b773936:0e805c3a	c1271fd3c8f2bf670ec91d0b:9dc5bd99fa98588067409f0e5fc6c68c:22d81ef474	M	t	9B
113	9a9908dbd674929ca3503bf5:3162e01a9645ce3a21d7494b8b704958:55b3391a	9237ca6510b00431d624b5e2:169d60b3a17a491badf248ef3e9284f7:5e717bab0e93f05b0d53	M	t	9B
2410	ec352b8e5953d594fb30ce8f:bf20d2d1a81e3c9d3371a1b15db83004:f6dde0cd	3671a8b2e0c171e8caf8c77e:67a2bb19e5f86838300d5ab0dc095a2d:7039bca162af	M	t	9B
1110	e203cb58f34d834f55bcce05:d3d2036acf2630d5c2ce7ebae4f0b479:46bfdce1	eeacea4f16130c2b01efa593:2bd3351d96a7653b9dfe816084c4ee0b:82e9896a8b8ad33c	M	t	9B
1944	f663ec659e9c6201e31e227b:4c8005bcd58566a472dee96fd176c1ee:4e257cebb2	d684942423f7d9165e9c0cc8:bf4096ed9f955868a5e8b66bc238dbf9:fca8887d4d	F	t	9B
77	6fc87431d36e806ac9204b29:fb53387ac0c940d14ecb6f9828143875:ca8e80f1	6bf19499353d8cc9926f429e:1ec96abc32c59491a4bf95d1e2fbea23:f043b42c7a49793a24	F	t	9B
1560	e3b01ced646568c3c4d85113:1f17ac28239bd799e2ad8ca15ba9f9f9:65731a8d2350	d2bc832ece28d724dd4b2b3e:ce54a10ff7d99b84d598d3effb84d7c0:dabcf182309f	F	t	9B
82	3201c579bcfaf7e06636ae0c:23acfcbb667b208f71ed8c85b061363e:b50e3830d5	6ee6f89c7876e0285533752f:f7c4e64ff08b721d33d45dbbe56b74cf:f7c439d445d4	F	t	9B
2604	def77f490ae2774020823780:b9a67b0599d01b580cac88ad0e1d1f8b:c9aeea61	cd8043cb9e7489099f05ba3e:20d3160e7a86c9db97bed5f5420a4c6e:e6192a4b85b55698d6	M	t	9B
1197	321397eec81085338a11dfd8:54af61ef1f55273aeb999246c299d7fa:9f6ebb70	d4859ca2ed024d45a3b82851:1fbf3387fbbcd3f8129e7ecce61a11b0:57d4e5471911	F	t	9B
2211	958c518e942b55d3c00609f0:7e26821d59f671bd6c1f2e6f837efe0c:606375ddec2960eb0c4a2e	e761341bf316bbd21cf06fc1:62e5f229889ba0e3a3c95963cb925f82:d5c5849751750167	F	t	9B
2429	398b48c10fb422ea37f78dcd:c2e640c8eefd947250ec852dc1520237:d9ac456513df0a261ebd10c5d3a4	8359c1b11e5efb2845acfa01:784980d573e2d11115352c9e137b28f9:dc1c71d4	M	t	9B
1460	875b0cdb2a52319aace557f0:2b0671f562eefde87ec32b44eebafa3e:baecd4bc0e7e	9f43208fb1b18ed84df95e05:993bb9f89ba15cee4ee510fcf0909786:8e37f85f7f	F	t	9B
1717	e358c6fb98ce384391a4b732:94e9551cc36eb0164af560d720ee71e5:0758dd7abb	c425d1cf736188ca0b695dd0:1a4bf035032fb8678e67c5737afdf988:d9b68436add41b	F	t	9B
2578	03511f471994f36628f54fe2:5d7a7fab5e7fab8e4ea8d5a2f1b1701e:0c84359b	9f0a35da8c74309c743fdbb4:a8b64182ac812890f414c89670d5d2f6:9a416fb85e12	M	t	9B
1572	d25ed6c4282f1d80995436c2:925da0b98b51adc1113c53ccdb97b3e1:fab7bb365e4c9b72	8f43a5ea2f5424dd5c9a67ea:6c494e0748c847853deb8efae65154f0:5b139d04d9	M	t	9B
1806	7b76e362ab49ba71d5865c9d:a04eddce9286cd4d131de2c1a75b81b8:3af862deb93a	946875a74e562cc98058288b:dce46fa09352f88117b65ee8e1688958:d0f7deed66f783	F	t	9B
115	dc976ffa0d2e9da36d23c1c3:f4671c54128c6628951e34da7bf5285f:9a88aa4f	f61069f2a46bdb33280e2843:b6b76c2d30c7850c1b75e4c05a58b77f:53c1ba868a	F	t	9B
2223	471328d4eb4ed7858e872703:97b72d81ef2527943fa805b3f4d6ed5a:500ce88c8805	a616d218ff48f1b4b1816d19:fa940ba572370924bd7bdc46c00f3a92:9de6f7fb4f0a17d88c	F	t	9C
1113	90045ed9f2d8562ff419f7bc:09bde2cdf3a929fba348d82a2de633e9:3f914f4f7d	c820b2729a7de90cd082425e:bdf38d280708840c7470956c951825bf:5861af60580f9c6b13	M	t	9C
2574	f4aa5d1986e0f8be87d6100b:e848255bb4b09ae84d4d66909b34430f:058d7241	7eee781d064b54439c0695b4:b65e6be984deb131afff48de8aa5e9af:ff854f18099ea87190b983	F	t	9C
1709	a9d8b3ffd3a27a4f2c57f300:f7102641ec546f6960f4073692b3be53:bb09d9d3974b	843e4d934f773cf5e1ba5e05:9d56da1ab41233e3cea627bd7270b02c:9cb1516e6851ceaddd	F	t	9C
1710	bd510ac35cdd5ab4dbff41fd:6a2f197a073edf532555c5e8006ecdba:996720771e40	0a3817534068cf62cf040f77:8ca3f4509414eea31768c501aef4d578:17bd2e	F	t	9C
1908	43e683418842ee7b11e63d18:7b16523dc2ef4ce4b38a8c33a9522b06:a69ae44122	d3847f2fd44f8dcc81d97637:65c0c1e3c1489c84912b790ab62e88f3:74bb62dc92b8222b694035a0	F	t	9C
2293	51f5f89eeea2fa31a3e8748f:5ec3371dd71b5bdd367ee2680df6b42a:e11d4d41	7f09c254a26c506e29a3b56f:066f49910a63c534eb21af5a5d6b504d:8c5e544ea8	M	t	9C
1905	fafa63842047b894777fa20c:69823294c9be4bd7b20a9534a40b76bf:3a2fd72c	c0d4c8bcd54b552562099fbb:c3331b89b188fb0b8346acbd94dbc286:ba15ad75	F	t	9C
2403	7f1a43e33a00cbd99f9de8b1:68dd5deb335ed03085880c66009aa208:df6d9f71e7fd95f5	42af8032fd0835e106c9089b:a5f86e9b5a900b96332cd02e4fcfd769:3b0d0f40366f	F	t	9C
84	e8b647cc974b93f6fceab4bb:2436731878e630f2c84a30f874219120:c430281efb	8bfe13fd8bd8cb944ae6e4bd:9521471e1f8c8e8d0d23c8ff04e53e77:3d391f97fb99	M	t	9C
2575	b2c1522f9b15904be28b1ca8:134cdd0bc49dd57fafab501cd62fb377:fc10ce57	1903ea2764e5a23933ad0a50:f86600f944f00edb11ce772f22878583:db64755544f37b	F	t	9C
93	16ea72b2a41401d97b3d588c:394eecb67189013eb3dd389ccdf6688c:91f3c12f91	c5949f389d258fabfde66d81:26e264c4818becbf02a3327494dac01c:0de743b4	M	t	9C
2219	39ea5f8db6f3f1d6523cc2d2:40a3d7e618c7c6dfb951b67ccd1fec08:ae81a577	a79c94287f9b28da6af93023:31bc6b5578a3ac3b5f14efc209cff77f:290b2f5f360efa	F	t	9C
782	afa6ebefc92e074e0cd58aa3:cc077f2c4cc67b124949f7b18b5315a7:b913a806c421	e8afd1237bed0e056e6f94d7:0b47758bebda0253afdf1ee01ef487d1:c9ac99323bf44439	M	t	9C
1539	a70fd2234ead2b0cb500b378:8291affdc331d8c3932c2c50a25be8bf:2f3a8fa5608b	15ff45d3456f094cb61c3b68:e00bd6b892d0363c1d5aa778051426bf:b036aa4111d56c	F	t	9C
2579	6a844ebb843e0af3a6593c62:bbec8f92113daa489f1b3e13d7116b0c:2933520040	b04df20e20fb1037b1b733c3:8fc4fc38752ad786e41313004e81b255:59702b162d	M	t	9C
2347	2d9fca50e373b8ea4b1a3b03:5f910a403a35343dca24c9e4fbc29fb2:7063340a7c316a	3d541a8499670a103d7006c6:2b5d2309b76edd7b56d84fd37b474a26:cb434a20	F	t	9C
2146	f33f249c7b6c29a644e100a5:7c97e4cf5a40a85cbd7a8068342d7148:85c6e0cb1d718a18	276021d48f6c17150c16ad16:276592048f891260b1e91eb0473f591a:dfe1b876af1fcd	F	t	9C
2029	ccd08b790751280529478068:c31eab31c49ccdd8b46ff244c086acf3:ef63ec54c1f5	3dfca5eb63c5c6aabbadf570:a4e965a1d5e27d9b6ca08747c33765a7:a42c906204	F	t	9D
1708	087ee5fb8c6d3b0d514cf0c2:a356cb991d960ed4048c76b6e1ba0dc2:f51a0d5de9e7	79ca5a485025245a2515fcc8:c2e495049f0217c35a193dab58a73a31:9ebaab467b8afd7411	M	t	9D
2142	7ed9e89d6fe4b83885c2a4de:529591c0cc9fa67797ba0714a401efc2:42809cff40	594445d6fa90b4ee25ce4d08:89ffff5e76e458a292b8ed05e0175023:644fa4f5bce230	F	t	9D
959	f1743507b813ad29f50259ac:961762f04c83aac0a3d8bf26b65c6e8e:44332e8a	60212ca5dba6ccdf09064808:70396fd69c853038dd3b194ac94eb049:431b58cfc5	M	t	9D
1811	7cef2f26b8fccc55d7fc74ec:4ef80221aa008515ac1a07d2982e9503:061f3f66	32df64966da9573d93689250:9418abe8f61f6c7207cbce582d4d6874:e71a070f59b93b	F	t	9D
2226	fd8d35291e7426f9e02d7e04:115b322cf13a726186eed37d691fce6a:f83a0285a9dd	1a2a60462f2b3be98d0a1b23:1eb382bd1cf71ce4fe7ba693fc0fffdb:e062118c61c3c289	F	t	9D
74	2207ef72b790927f638ff9c7:21c0a43747219722ac18166edbfe0998:65004f23550c7c6a	7fecc9fdc00db3ed2052525f:710893c73f4a64975a09e2f8e3435ac2:ed8fba282b1036	F	t	9D
1188	773a0fe4e04ac989a84913bd:2e3fd7e8e17ff5a95db3c79b288f51f2:db3bb8e7d6f1	d59304904785afb36bb911b2:6e048effa61030bdf9c2ecd772b73e2b:02df46cda633fa61cf7109b1	F	t	9D
1839	a49fa9afad041562895239c1:def8b9fad09718bd8fe01e26c9d20873:aa27083d0e32377ae1	3c6b8e2f0ed038834b266a0c:164eee6b2c427feae2e395cd49419fa5:2ea18d266c0356a127674f33edc9	M	t	9D
1715	672630eb5137cc55200faece:f61118b7cdd05d7d67277670dccb1d2c:420b03bcd07687b07c7a	b7a9460e0da16a62d932da74:0481b3803de5f99d699fd3f8bc522918:9e4442dbfd2ba6	F	t	9D
1189	98f7d76c8dfa0b8bfacb15ee:69a08dccd0da89174ff9d1af3915071a:45e9e0726a72dbc0	793815cdbd576d314314154e:5a342b694b883cadd6364767ea2dc3de:0b2f566bea1c	M	t	9D
94	3da3539849b142a796c8a144:ea39e94d3498b0d64ee5f4bf37a458f3:6e695b7b33	edc85e9a19f4ffa4f7f17158:03f32402ae588983b0e5e9ef3f02e466:96f3a2b4	F	t	9D
97	83dd5dff132005ddc7f73116:c1991c03dc9768778ba1a90028b3bb15:4524cca8	facdcbf490efa27ec2f45b25:2a74788c93f6e6e70805e0abfda6e268:6e0eaae78c41	F	t	9D
1193	872080c7a693c5609fa38f8d:1c1c415faece690693e7decc0d7b29a3:25b30d6d	127a0769460210c33805972a:3810403f07783bc4598506149af79d58:be37a89d58f0	F	t	9D
102	883bc351d31f0ac312c2be00:fc2b42cbec7313af5c2ddd4e821412e3:cdece78d7ff70019d6df8eb5	d7f61e0f5e3a67192bfae610:e1377175ec20b3f4cbb6cbe7f2b8096f:1f42ab6d8ea5	M	t	9D
969	bfd2e620856f4aee64d016cd:37d2034822fc4b9da9ef3fb77343bffb:68a9825fd176029901	842cbbb6f75a0b0b2a793aae:d80631d6b235a12bde15e88c03443aa4:bd02c3dc41	M	t	9D
1133	c56b0d24eaac5e1a01e90a8c:451b80e5d2807820deb247b3729e0158:6c125a4b9a253d89e03c0ce674d7	99c13ad1c4c9b82b68d32811:8e481e3d357da5bc479da7e7473f36cd:d1bbd1518ac7b6f8ed610f9d	F	t	9D
110	d5a6f9e7991cd1c9fbf43388:ae9c7b3bac308d6ef80eb20b836ba22f:4260e1b7950c9952	411ee4bc6d7cd866c197b0e9:e329a887657a4aa0e4efba4212089a51:36f3ae77a6ed	M	t	9D
1830	a7276e22cefb7ece43d3c6b3:1326cfbb9585f66d9a17ddba7b7449f6:c1e6d8ef8d85c7	9c1480201a938187689a6f78:b516b73e88100b4a08d1418d6270d426:77aefc278d62eec8	M	t	9D
1592	4d6b4bebb587d9b2024982f4:2f4e2eeda52db2159332f3e9bef7608c:329e816510	a14188be16d8ac46016c0d5f:7216f993f23cd82f6faa64396c7e18aa:5994f243cb	M	t	9D
1721	Ahmad	Abu Shamma	M	t	10A
1542	Jorge	Dominguez Narvaez	M	t	10A
2304	Aya	Elabid	F	t	10A
2149	Sara	Elsayed	F	t	10A
2154	Yahya	Kamran	M	t	10A
2258	Azlan	Khalid	M	t	10A
169	Abdulrahman	Mahgoub	M	t	10A
2253	Alleyah	Mahmud	F	t	10A
2335	Adham Negmeldin Mahmoud Hassan	Mohamed	M	t	10A
173	Elma	Monzer	F	t	10A
175	Lellian	Nassar	F	t	10A
186	Hossein	Rakha	M	t	10A
2178	Joana	Rifat	F	t	10A
2180	Ayisha	Shanta	F	t	10A
2157	Fatimah	Shareef	F	t	10A
1382	Lucia	Vargas	F	t	10A
199	Reeda	Yafawi	F	t	10A
1949	Maher	Zain	M	t	10A
116	Adam	Abdul Hadi	M	t	10B
119	Joury Rami M. N.	Abuzaid	F	t	10B
2148	Malek Mostafa Ahmed Mostafa	Afifi	M	t	10B
2298	Amanda	Aguirre Cardenas	F	t	10B
1087	Aliza	Akram	F	t	10B
129	James	Bejjani	M	t	10B
1929	Emily	Boom	F	t	10B
2235	Muhammad	Faisal	M	t	10B
2030	Agustina	Flores Ibarra	F	t	10B
1723	Riccardo	Gennari	M	t	10B
1776	Maryam	Ghaly	F	t	10B
1208	Simrah	Iqbal	F	t	10B
1011	Woojin	Kim	M	t	10B
174	Hassaan	Mudassar	M	t	10B
1724	Ishaq	Nazim-Uddin	M	t	10B
177	Minha	Nureena	F	t	10B
181	Maximilian	Paddock	M	t	10B
2343	Raihan	Paramban	M	t	10B
183	Levent	Pekey	M	t	10B
2376	Rebecca	Van Der Horst	F	t	10B
2005	Mahad	Zaman	M	t	10B
1935	Aya	Al Wathaifi	F	t	10C
1911	Ziad	Awadallah	M	t	10C
130	Tristan	Benfield	M	t	10C
131	Aisha	Bhat	F	t	10C
1541	Sofia	Cipriano	F	t	10C
138	Omar	Dasouki	M	t	10C
2301	Samaira	Diwan	F	t	10C
903	Tayim	Feghia	M	t	10C
2152	Hala	Hegab	F	t	10C
2153	Santiago	Herrera Rubio	M	t	10C
162	Haniya	Khan	F	t	10C
2616	Lorenzo	Matos Valero	M	t	10C
1817	Mohammed-Adem	Menacer	M	t	10C
2156	Anayah	Noor	F	t	10C
2242	Luana	Peirano	F	t	10C
188	Ainoor	Rauf	F	t	10C
189	Hiba	Rizwan	F	t	10C
2158	Ayush	Singh	M	t	10C
2224	Ahmed	Akmal	M	t	11A
1020	Maryam	Alghamdi	F	t	11A
314	Islam	Dabjan	F	t	11A
2214	Ibrahim Omar	Eid	M	t	11A
318	Kareem	Elkhamry	M	t	11A
944	Abdullah	Faisal	M	t	11A
2162	Yahya Ayman	Hegab	M	t	11A
1802	Abdullah	Kamran	M	t	11A
1022	Mohammad	Khawaja	M	t	11A
970	Vishal	Lakshmikanthan	M	t	11A
2217	Anwar Ibrahiem	Mohamed	M	t	11A
2365	Yasmin	Mollayeva	F	t	11A
1915	Mia	Mourad	F	t	11A
1999	Alikhan	Nurlybay	M	t	11A
1963	Muhammad	Sahm	M	t	11A
343	Kumail Muhammad	Saqib	M	t	11A
2181	Yassin	Soua	M	t	11A
1919	Mohammed Ramin	Valiyakath Karakkad	M	t	11A
1210	Tianyi	Wang	M	t	11A
1770	Mohammad Abdullah	Ahsan	M	t	11B
2160	Maryam Ibrahim	Akbarzade	F	t	11B
1726	Omar Mohammed Gamal	Aly	M	t	11B
1727	Abdul Rahman Bin	Azhar	M	t	11B
730	Racimbey	Bendjaballah	M	t	11B
2363	Hazriel Aqim	Bin Hafis	M	t	11B
2245	Ethan Cheuk-Han	Choi	M	t	11B
319	Omar	Elleathy	M	t	11B
2161	Malak Ahmed Elsayed Abuelmaati	Elsayed	F	t	11B
1212	Keltoum	Farhi	F	t	11B
2239	Jude	Khatimi	F	t	11B
329	Rayan Fawzi	Kodeih	M	t	11B
302	Khalid	Al Majzob	M	t	11C
306	Abdallah	Al-Arakzeh	M	t	11C
307	Naya	Al-Damluji	F	t	11C
1914	Anaan	Anver Sadath	F	t	11C
316	Ahd	Eldessouky	F	t	11C
2394	Annie	Erebor	F	t	11C
1840	Luis	Flores Wiedler	M	t	11C
1213	Jude Mustafa	Hajjar	F	t	11C
323	Shayan	Jamal	M	t	11C
2172	Lucia	Milsom	F	t	11C
1730	Muhammad Irfahan Bin	Mohd Iskandar	M	t	11C
958	Hala Eslam	Radwan	F	t	11C
2165	Drudzaki	Tafsiri	M	t	11C
2182	Fatima	Valiyeva	F	t	11C
1610	Junsang	Yoo	M	t	11C
982	Hesham	Abuanzeh	M	t	12A
1930	Kais	Al Wathaifi	M	t	12A
363	Yahia Mustapha	Berkane	M	t	12A
2580	Hayah	Faizan	F	t	12A
1217	Minha	Furqan	F	t	12A
1466	Yasseen Waeel Ossama Sayed Ahmed	Hamouda	M	t	12A
2166	Mariam Adel Galal Abdelmordy	Helal	F	t	12A
375	Loay	Hilal	M	t	12A
380	Hamdan Raza	Khan	M	t	12A
384	Talya Saleem	Khashan	F	t	12A
2589	Yejin	Kim	F	t	12A
1761	Enis Kayra	Koyuncu	M	t	12A
387	Selena	Mahmoud	F	t	12A
2167	Te	Ma	F	t	12A
393	Kyera-Zoe	Moussa	F	t	12A
1545	Bahadir	Ornekci	M	t	12A
2603	Abdul Ghani	Saoudi	M	t	12A
2590	Isra	Souiai	F	t	12A
1736	Braxton David	Wright	M	t	12A
1418	Shayaan Ainul	Abedeen	M	t	12B
1732	Madhawi Abdullah A	Alotaibi	F	t	12B
361	Yusuf Tofek	Bahomed	M	t	12B
1755	Ariana Abdalla	Bazzari	F	t	12B
374	Abdullah	Bin Hassan	M	t	12B
1227	Alia Ahmed Farrag Abbas Mahmoud	Hafez	F	t	12B
1222	Sanaa Fatima	Iqbal	F	t	12B
383	Zarah	Khan	F	t	12B
909	Areesha Moazzam	Khawaja	F	t	12B
386	Dan Mohamed Ibrahim	Mahgoub	F	t	12B
2168	Seif Islam Mohamed Safwat Mohamed	Maklad	M	t	12B
395	Zoha	Mudassar	F	t	12B
396	Emad Adnan	Nassereddine	M	t	12B
1467	Eesa	Nazim-Uddin	M	t	12B
398	Omar Moutaz	Osman	M	t	12B
1735	Alessandra Kartika	Pijoh	F	t	12B
402	Mahnoor	Ruheena	F	t	12B
413	Bahjat	Wadi	M	t	12B
1737	Brylie Alyson	Wright	F	t	12B
1974	Lamia	Aleem Siddiqui	F	t	13A
880	Samya Mohammed	Almulhem	F	t	13A
735	Martin Shady Wagdy Farag	Asaad	M	t	13A
427	Sophie	Bejjani	F	t	13A
737	Dana Louisa	Benfield	F	t	13A
436	Zainab Mahmood	Din	F	t	13A
981	Maya	El Malki	F	t	13A
460	Joyce Adonis	Monzer	F	t	13A
523	Ali Eqan	Saqib	M	t	13A
468	Eshal	Saqib	F	t	13A
1472	Muhammad Hamza	Sarfaraz	M	t	13A
1794	Matvey Andreyevich	Spencer-Sokirzhinskiy	M	t	13A
473	Anan	Wadi	M	t	13A
474	Abdullah	Zeeshan	M	t	13A
420	Shamekh Mohammed	Al-Arakzeh	M	t	13B
1738	Noor Osama	Alburaiki	F	t	13B
2292	Roaa	Aly	F	t	13B
1926	Tamara	Awadallah	F	t	13B
1598	Aiden Mahan	Dindial	M	t	13B
448	Orla	Hannan	F	t	13B
1471	Hanan	Kharbotli	F	t	13B
1928	Reine	Mourad	F	t	13B
1561	Ioannis	Theodorou	M	t	13B
1184	Hanxue	Wang	F	t	13B
976	Alexander	Zaour	M	t	13B
\.


--
-- Data for Name: PupilCode; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."PupilCode" ("id", "assignmentId", "groupId", "code") FROM stdin;
\.


--
-- Data for Name: Role; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."Role" ("id", "name") FROM stdin;
role_admin_001	admin
role_hod_001	hod
role_teacher_001	teacher
\.


--
-- Data for Name: Subject; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."Subject" ("id", "code", "title", "studiedComment", "commentFormat") FROM stdin;
DT	KS3-DT	KS3 Design and Technology	\N	<Communication>\n\n<Make Skills>\n\n<Materials And processes>
b48a4d58a4743bd8b2db	AR1	Arabic 1st Language	\N	\N
8c692847d52879b4be83	AR2	Arabic 2nd Language	\N	\N
c8c86430aabc681813e8	ART	Art	\N	\N
c5603dd1210b371db7ae	BIO	Biology	\N	\N
0f4de0c5ad45017e88f5	BTECBUS	BTEC Business	\N	\N
4181c022d9eb05117c6d	BUS	Business Studies	\N	\N
b148232e8ddf91bce214	CHEM	Chemistry	\N	\N
6196f2f5279ed1941ab7	COOK	Cooking and Nutrition	\N	\N
c8d7bfa257934576575e	CS	Computer Science	\N	\N
abff90061fc53561024e	DM	Decision Maths	\N	\N
604ab32c3290fb9c3a31	DRAMA	Drama	\N	\N
bf43ccb47927cefba4e8	DSCI	Dual Science	\N	\N
dfe36bbe8755cfb5ba25	ECON	Economics	\N	\N
dfa5c7856b40dd6c8c80	ENG	English	\N	\N
5eca2fa95cdf368d7fba	ENT	Enterprise	\N	\N
34b041927ed21b2b5cae	FM	Further Maths	\N	\N
45c516b9a78910b2536b	FR	French	\N	\N
8d14ff7e9d97595155c0	GCSEPE	GCSE PE	\N	\N
d4b8fb5671d44379b6ca	GCZ	Global Citizen	\N	\N
d0c22c9fbde4eae78a2c	GEO	Geography	\N	\N
f2b13fe00258ff012e2d	HIS	History	\N	\N
6d30ecf1c96123a08f97	ICT	Computing	\N	\N
bd0c9a5ae8c3c14861e7	MATHS	Maths	\N	\N
4096cb88b43e598d2a72	MECH	Mechanics	\N	\N
d3ca3143b0f1c932737b	MUS	Music	\N	\N
194bd4628e9e1172ea99	PE	Physical Education	\N	\N
450588febb5ff75e8d45	PHY	Physics	\N	\N
8fbfd22f3e60c619e941	POL	Politics	\N	\N
647cdbe0df378408fdef	PSHE	PSHE	\N	\N
0740b4645b4357d8dfed	PSY	Psychology	\N	\N
888ccb16fd8e7504cc67	REG	Registration	\N	\N
2d060f26699b06743f31	SCI	Science	\N	\N
50c923e9e7b41a535325	SOC	Sociology	\N	\N
ff1729d8c7e693c5b111	SPA	Spanish	\N	\N
b1dd6ef556c9ec20224b	STATS	Statistics	\N	\N
\.


--
-- Data for Name: User; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."User" ("id", "username", "password", "isActive") FROM stdin;
user_admin_001	admin	$2b$10$71p.eB.Wd9t8XfEoWwL5F.iK6FzDapMVYeHUcDvWbamZ2HWRH/6ha	t
v152mbwu3vdblom8o3xwoad5	leroysalih	$2b$12$7tcEJoP1lC8xwJTCtnAl3efQ3Ml.itPhT7Dum8tXIsNuzY1sdYea2	t
7df17fc3ee9767f27492	a.evans	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
9b6a3472cedfe4f10826	a.kheralla	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
1b98a78cee88b43a5024	b.reese	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
f150e71801bd8410d574	c.henderson	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
f1a15047dcdc49b0d1bf	c.kapumpa	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
b09ff1ab1d0583aa955e	d.meadwell	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
5f95f3434be5648b5d80	e.annoni	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
f3a9ac717be348cb8932	e.gorman	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
513de73970f569aa0344	e.kane	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
a861497b23a650f9b440	e.low	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
7f501fe35061978683a9	f.berouijel	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
6298d554d8e14067ff3e	h.cox	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
e16f16d7233090cb7965	i.assaf	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
c89abddc6d32037da55d	i.henderson	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
dc3ba9a20e9d40aff0b4	i.kaul	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
d03263d74be35d68b7e7	j.costelloe	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
20e179146eb6bd982565	k.cunningham	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
e3c5fa508ad4fe4f19d6	l.ashton	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
5b9f00952b7726608a95	l.edwards	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
eb100bda8b0d3ecc0c62	l.gouk	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
7127cdaa6d5aac38b1ab	l.salih	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
2e42bbb99dd33754220a	l.talbot	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
0f2cb4f68319568db9ac	m.abubakr	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
6e6ca8c1794636815991	m.oneill	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
0aaeed216fcaa328c260	m.tyre	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
32ee1105401237bff687	m.yadollahighadikolaei	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
72b3ecb6bb3ddf696ffc	n.cook	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
3ab4e5dd7f5007d04f91	n.nash	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
11a6a61bf1f0c25b9e92	n.yousaf	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
ba907005863effc8887e	o.nash	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
77fe10b9e66e0427adda	o.sodeinde	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
e5bd8960caeadda4ba86	r.brownafari	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
4dcdd75becac3c5ac374	r.husher	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
26cbf254ce12996ba47a	s.barton	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
b4b61e4d64204228f13a	s.organ	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
fd0e10eec0a7cb76bd4f	s.rigby	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
e9913bf2396436bfd948	s.smith	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
158d15e754e983031efe	t.yafai	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
fd1e58160c741b9cbf3a	y.dasouki	$2b$10$pmrK1C4WWawcCLmRHCzrXeZv3PISyHOl8OZ0Nnqcd8CNCCQB84x5.	t
\.


--
-- Data for Name: _ClassToUser; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."_ClassToUser" ("A", "B") FROM stdin;
528f7d8434d6f3a6f7ad	513de73970f569aa0344
5315b384ce058e8ab8a6	dc3ba9a20e9d40aff0b4
283a8218e792eb7044a2	72b3ecb6bb3ddf696ffc
e48d91320929058cd529	f150e71801bd8410d574
e48d91320929058cd529	f3a9ac717be348cb8932
8082e012c8ef47aeca2d	f150e71801bd8410d574
8082e012c8ef47aeca2d	f3a9ac717be348cb8932
bd14ad4dd640660f1536	72b3ecb6bb3ddf696ffc
fc3cc49d60e02ea77ee2	dc3ba9a20e9d40aff0b4
0e647084d6478b6ffe86	20e179146eb6bd982565
3f5084d55890068aef0f	3ab4e5dd7f5007d04f91
c44e345dd4e716d24329	5b9f00952b7726608a95
fda056455d00bf047c21	9b6a3472cedfe4f10826
159a5eac4538c61ab219	11a6a61bf1f0c25b9e92
848c20d72e2e199b8639	7127cdaa6d5aac38b1ab
848c20d72e2e199b8639	3ab4e5dd7f5007d04f91
6d7d4c292566a7ad40ef	6298d554d8e14067ff3e
24f8f79de7d8daf5cdb1	6e6ca8c1794636815991
69a272fe2ae16f6f7a3d	b4b61e4d64204228f13a
4b02a403546e54cc7f5b	2e42bbb99dd33754220a
62216e3c0b97d44041eb	d03263d74be35d68b7e7
d0011180a38970fb74e7	0aaeed216fcaa328c260
51b76e207ca154e8acba	158d15e754e983031efe
3e4a03a8fd0bb2afc961	1b98a78cee88b43a5024
c5a5a84cddf882001a60	0f2cb4f68319568db9ac
dd1b0f7f6f3aefb5955c	a861497b23a650f9b440
d7e3398ffa52ca717c5c	f1a15047dcdc49b0d1bf
ebee19666521dd14842b	77fe10b9e66e0427adda
d6bb5024d5ec592d3c44	e9913bf2396436bfd948
c26810bb7ffda87028cd	eb100bda8b0d3ecc0c62
2e959c074f854901899e	3ab4e5dd7f5007d04f91
b4234235efeea20a0ca7	fd0e10eec0a7cb76bd4f
fd0e3630e13db624ea01	fd0e10eec0a7cb76bd4f
a0eea0aac26bc2c9a33b	9b6a3472cedfe4f10826
ec74201ba1740e7d0f91	e3c5fa508ad4fe4f19d6
24076f2e2af052db9360	e9913bf2396436bfd948
5f5846cdcdc67e6b6535	6298d554d8e14067ff3e
81ceff273e6f7f288ad5	dc3ba9a20e9d40aff0b4
d8a47a56012378b52d48	e16f16d7233090cb7965
c6854f08ccb4b2c8da4d	513de73970f569aa0344
98cc7263651f31776990	dc3ba9a20e9d40aff0b4
dfb67ad547da1c74dec0	72b3ecb6bb3ddf696ffc
366c6f405b1f8bc78439	f150e71801bd8410d574
cbbb379c9906538aaaea	f3a9ac717be348cb8932
8ce3f350dfe12ca44f95	11a6a61bf1f0c25b9e92
08cda36fbe812c092c9c	3ab4e5dd7f5007d04f91
53889c63251a67f47a2c	20e179146eb6bd982565
4b963bec52944fc95745	9b6a3472cedfe4f10826
6189d54b165dab8bd56e	9b6a3472cedfe4f10826
a82276e194b1732711b4	dc3ba9a20e9d40aff0b4
af2892f8639e42fdc83d	3ab4e5dd7f5007d04f91
02938d7b62625b948141	7127cdaa6d5aac38b1ab
f27cf82c1214ada6fb4e	158d15e754e983031efe
d131a3b33401a7b4a668	6e6ca8c1794636815991
ecf47a0bf5f29be17314	4dcdd75becac3c5ac374
ecf47a0bf5f29be17314	b4b61e4d64204228f13a
1b29574d86db3b63e11e	6298d554d8e14067ff3e
0bfa86e7a44f47e58019	2e42bbb99dd33754220a
0bfa86e7a44f47e58019	b4b61e4d64204228f13a
d13bbcf819ba7aedcab8	fd1e58160c741b9cbf3a
4dbde706728519fe11b9	0aaeed216fcaa328c260
4dbde706728519fe11b9	ba907005863effc8887e
e295ccde749fd2982eeb	7f501fe35061978683a9
ffc06f66d5b1395c4fe6	158d15e754e983031efe
96c9cb3e39287f99c251	7df17fc3ee9767f27492
96c9cb3e39287f99c251	0f2cb4f68319568db9ac
9035d56968d798f556cb	32ee1105401237bff687
ff8a7abaa5467496e924	a861497b23a650f9b440
ede2ca1d19344a1b8138	1b98a78cee88b43a5024
1ef93a4c890e726a2aca	77fe10b9e66e0427adda
15341f322b282b495094	e3c5fa508ad4fe4f19d6
2a8c35f329bd530acc98	e3c5fa508ad4fe4f19d6
7788942930b1d0ff75cf	e9913bf2396436bfd948
8ecde88989703bbe399b	eb100bda8b0d3ecc0c62
f212ea07533e401eb4df	11a6a61bf1f0c25b9e92
591c7cad2de417c3536e	fd0e10eec0a7cb76bd4f
9ea43fbf389f497e0a95	fd0e10eec0a7cb76bd4f
f3f87a4f6848eef065be	e5bd8960caeadda4ba86
fbd0f00af5ce74d958eb	6e6ca8c1794636815991
a89ce690562f1bfb0bba	f150e71801bd8410d574
a89ce690562f1bfb0bba	ba907005863effc8887e
6e57422643f6d74880c0	32ee1105401237bff687
0369837b7f5d361afb51	513de73970f569aa0344
d664df57dc8d387d1886	f150e71801bd8410d574
50a78e0f85245df78ebb	3ab4e5dd7f5007d04f91
2399822855f02dd17bd4	72b3ecb6bb3ddf696ffc
2399822855f02dd17bd4	3ab4e5dd7f5007d04f91
f30291eeab3f00e3032f	f150e71801bd8410d574
f30291eeab3f00e3032f	f3a9ac717be348cb8932
7c1a4815aa1a2d69e547	eb100bda8b0d3ecc0c62
46d83744c63768dc907e	11a6a61bf1f0c25b9e92
0b72791b2dd715141534	a861497b23a650f9b440
29e5c9ab72aa19cf5199	6298d554d8e14067ff3e
29e5c9ab72aa19cf5199	b4b61e4d64204228f13a
242ae56440681672b491	f1a15047dcdc49b0d1bf
af736540ce33e44663db	7df17fc3ee9767f27492
86f346656a8fafa3495b	0f2cb4f68319568db9ac
ebfcbb48fbd63e18339e	f1a15047dcdc49b0d1bf
7f1e9e42b1894464421d	a861497b23a650f9b440
a5012b7bbf099fe29fe8	1b98a78cee88b43a5024
f6dcd719304f00d82c2e	eb100bda8b0d3ecc0c62
bb68063696175471a537	fd0e10eec0a7cb76bd4f
a72e2dd120ef2c86ef04	d03263d74be35d68b7e7
7e2d798388bf048f20f6	e5bd8960caeadda4ba86
1b092bfce595f85293ee	f1a15047dcdc49b0d1bf
b220f2ecaea0a0fb00dc	c89abddc6d32037da55d
6e2aeb5d8592447a31d6	513de73970f569aa0344
eb6a091792f47c508880	dc3ba9a20e9d40aff0b4
c3c9a4fcec3d5a4cc8f5	f150e71801bd8410d574
c3c9a4fcec3d5a4cc8f5	f3a9ac717be348cb8932
c7844f6432c9b3505784	eb100bda8b0d3ecc0c62
e6e6b9eaef349d8f67bb	7df17fc3ee9767f27492
5d0f2a46fc3857533f25	4dcdd75becac3c5ac374
5d0f2a46fc3857533f25	b4b61e4d64204228f13a
b29c20c00222a6c1876c	5f95f3434be5648b5d80
b29c20c00222a6c1876c	fd1e58160c741b9cbf3a
75a0ab6961da49ab3b89	0aaeed216fcaa328c260
2aba875ca65d820f85b4	7df17fc3ee9767f27492
390155dc1ec0a20487bd	32ee1105401237bff687
30720b1ba7b7ea205543	9b6a3472cedfe4f10826
00a45de906cbd0942534	fd0e10eec0a7cb76bd4f
f75e70904a70afafbbd1	d03263d74be35d68b7e7
552ab552eb8762c570a0	32ee1105401237bff687
78a0f435af1bd5acb2c7	f3a9ac717be348cb8932
9bbe8ed0f5c6182ff324	2e42bbb99dd33754220a
6f69f172e874d825216c	e16f16d7233090cb7965
40e35d563990a350b589	fd1e58160c741b9cbf3a
ade303ded2be972c6a5d	513de73970f569aa0344
37467e477610598c1474	b09ff1ab1d0583aa955e
ce89e455bbdc905dd759	7127cdaa6d5aac38b1ab
5e6582e3fe3b905468a7	5f95f3434be5648b5d80
98559a16949b4230e009	d03263d74be35d68b7e7
6b755b3363797e6feca4	e5bd8960caeadda4ba86
d13ffb488e82f76b02d0	20e179146eb6bd982565
310c982ca10f2d32ccb1	77fe10b9e66e0427adda
d421b774241434ccd836	e3c5fa508ad4fe4f19d6
39b1ab8e2df8c0182786	72b3ecb6bb3ddf696ffc
b630f79c61b93bb9a72b	72b3ecb6bb3ddf696ffc
3649e9e9359e028f0892	5f95f3434be5648b5d80
b6bbbaae703037f495be	72b3ecb6bb3ddf696ffc
0aeb1b79309469c06c74	4dcdd75becac3c5ac374
0aeb1b79309469c06c74	26cbf254ce12996ba47a
567b97f8b1de47a17d7d	6e6ca8c1794636815991
567b97f8b1de47a17d7d	26cbf254ce12996ba47a
a7de483e1b81ba1ede04	6298d554d8e14067ff3e
a7de483e1b81ba1ede04	26cbf254ce12996ba47a
0dae9224680dbe8d6673	a861497b23a650f9b440
492be8b28f568157cbef	f1a15047dcdc49b0d1bf
d3f8ea02cba4d989b5f4	1b98a78cee88b43a5024
89788763a94d40f933fc	32ee1105401237bff687
94b4fdf386152f607a13	2e42bbb99dd33754220a
94b4fdf386152f607a13	26cbf254ce12996ba47a
0f69da4484eb0767a720	e16f16d7233090cb7965
dc667f3c1235dee04b67	fd1e58160c741b9cbf3a
0e8355bb2b56b6ccda79	513de73970f569aa0344
96ead37b6d717ecf4c9d	b09ff1ab1d0583aa955e
dd84dd43e18cd13b299e	7127cdaa6d5aac38b1ab
7f4d00c651d1bc22f95e	5f95f3434be5648b5d80
be3719af94743bcb4e1b	d03263d74be35d68b7e7
c34e64cafe735b50ed2e	e5bd8960caeadda4ba86
c69b3760d1f706da6487	20e179146eb6bd982565
f82c00f0fb8ba676d79a	77fe10b9e66e0427adda
d5cde185f3020a8d2c79	e9913bf2396436bfd948
da2af99abb1da9ed1063	5b9f00952b7726608a95
ac8714c2edc7a6940ce7	9b6a3472cedfe4f10826
335b5ae685bab3b50f30	5f95f3434be5648b5d80
56061cc582e767291018	5b9f00952b7726608a95
0ab34b12507bd5ae0386	e16f16d7233090cb7965
09c52f1ac992e352a953	fd1e58160c741b9cbf3a
d15d7b9fd2ed7e46dd53	513de73970f569aa0344
859e00b64f3807defcfe	b09ff1ab1d0583aa955e
87ccbe2ce9a680ff93a0	7127cdaa6d5aac38b1ab
a477427b675c489d36f6	5f95f3434be5648b5d80
0e634788a37d82abbef4	e5bd8960caeadda4ba86
d7cef190d3651254a780	7f501fe35061978683a9
2a9db3e19b93ce9c355d	20e179146eb6bd982565
98d9b41d2eb6d4fc61c8	77fe10b9e66e0427adda
dbe35887465186683743	e3c5fa508ad4fe4f19d6
7e1c155a4ca800a30d86	5f95f3434be5648b5d80
8cdd523b41f1a026f4de	5b9f00952b7726608a95
061ddff82e7b822b7ca7	5f95f3434be5648b5d80
abe9d4482e5fa2c5fa0f	5f95f3434be5648b5d80
05504bdd46170dd3f4db	e16f16d7233090cb7965
83e93fed391058692f49	fd1e58160c741b9cbf3a
6db953d26eebb36848ad	513de73970f569aa0344
22eceb505d094f53c53e	b09ff1ab1d0583aa955e
a8ff0afc8a2eaa57d68b	7127cdaa6d5aac38b1ab
273f5b438cc3414ba44f	5f95f3434be5648b5d80
ad8ed2ff59798b5df7ba	0aaeed216fcaa328c260
ad8ed2ff59798b5df7ba	ba907005863effc8887e
8639aeda2bf449728d9d	e5bd8960caeadda4ba86
b4ccb16d04bc3b14647d	20e179146eb6bd982565
b4f371ad7826b1be4b8f	77fe10b9e66e0427adda
bc02a71806dababec353	e9913bf2396436bfd948
4d66ba7137d37d690550	0aaeed216fcaa328c260
d9c04b00fffcf9b34420	5b9f00952b7726608a95
011da44732c7f182de66	5f95f3434be5648b5d80
0af02ffbcd58d7253c76	0aaeed216fcaa328c260
15d8cb818f560a27d9a8	e16f16d7233090cb7965
ccd8ac5ce51d687a65b6	fd1e58160c741b9cbf3a
ccdd93e22ce8fde4aefa	513de73970f569aa0344
81f8ac02e568df5d9552	1b98a78cee88b43a5024
3ea70d2daf15198c70a6	b09ff1ab1d0583aa955e
4a894a4948ebc3d652ed	7127cdaa6d5aac38b1ab
62a6aa3edf4fe2dee2d0	ba907005863effc8887e
62a6aa3edf4fe2dee2d0	26cbf254ce12996ba47a
cd4dde106f1d53c9aada	5f95f3434be5648b5d80
cd4dde106f1d53c9aada	fd1e58160c741b9cbf3a
011f0cc85f1ad1cc22e6	0aaeed216fcaa328c260
be0802e0a01e96e75eab	7f501fe35061978683a9
2ee37f6a5b5afdba3728	20e179146eb6bd982565
042f404c813aed6ee038	77fe10b9e66e0427adda
9538a1857d657c0f004b	e3c5fa508ad4fe4f19d6
998e79a5a2df8f596fa3	a861497b23a650f9b440
bd6f2ca2ed437f9499d2	5b9f00952b7726608a95
d74ad31e1cd4b3e1a678	a861497b23a650f9b440
9b82d77926477514bec2	f1a15047dcdc49b0d1bf
dcfa4362bb9ba014e5b6	32ee1105401237bff687
70d175493126d8dceaf2	7df17fc3ee9767f27492
67cb84d4c26e5522b07c	0f2cb4f68319568db9ac
7d66aa886cfe3a72f86b	e16f16d7233090cb7965
d8bb7e8a8ebe34b835c3	fd1e58160c741b9cbf3a
1993cf9763443fce0f82	513de73970f569aa0344
4901021e0e8f2486bbf9	1b98a78cee88b43a5024
b1708239145b8854c8a3	b09ff1ab1d0583aa955e
17be688be816f0468b05	7127cdaa6d5aac38b1ab
6e3f730a3b6e866694e3	6298d554d8e14067ff3e
6e3f730a3b6e866694e3	26cbf254ce12996ba47a
7dc8596f01b7d5f27d2f	5f95f3434be5648b5d80
7dc8596f01b7d5f27d2f	fd1e58160c741b9cbf3a
58f247c4f27334049744	0aaeed216fcaa328c260
417a1d3104a891ccfa38	7f501fe35061978683a9
9a01025dd44aaf5de999	20e179146eb6bd982565
fc618b2ec9d2df99322d	77fe10b9e66e0427adda
52a0c7618351d4fe83e9	e9913bf2396436bfd948
332ec0589e88163cdbc2	fd0e10eec0a7cb76bd4f
c831500edcf62a572d96	9b6a3472cedfe4f10826
c831500edcf62a572d96	11a6a61bf1f0c25b9e92
942711e71cb8fb753a49	fd0e10eec0a7cb76bd4f
140cb8e7fd7f918244a8	e16f16d7233090cb7965
0f9cf77856a913b30c8e	fd1e58160c741b9cbf3a
67a7363447e757702674	513de73970f569aa0344
a09b6ca1c047b268f7cc	1b98a78cee88b43a5024
ffecd66e5100127c4042	b09ff1ab1d0583aa955e
6c05b760b3d8e03e72a7	7127cdaa6d5aac38b1ab
937d339e9ff57e16d3c9	2e42bbb99dd33754220a
dac65aeaaed135b88dba	5f95f3434be5648b5d80
dac65aeaaed135b88dba	fd1e58160c741b9cbf3a
4883160b03bc92b91640	ba907005863effc8887e
d87f1150ff27c1efbcc2	7f501fe35061978683a9
524cb475f7eec63c3736	158d15e754e983031efe
9b8fac85cf8101939f27	77fe10b9e66e0427adda
0716aa5bb459f01a8e41	e3c5fa508ad4fe4f19d6
fe2dc648d4dc41df5d30	e9913bf2396436bfd948
42c2357c9e79a8d2bd9e	11a6a61bf1f0c25b9e92
42c2357c9e79a8d2bd9e	72b3ecb6bb3ddf696ffc
1bfba56a05f5800f2691	e9913bf2396436bfd948
c38a7595c8fe93bb3a4e	e16f16d7233090cb7965
8b094bf3bca9e805dbab	fd1e58160c741b9cbf3a
922a219336536e10a18d	513de73970f569aa0344
5e3639a7c71b3a997aa8	1b98a78cee88b43a5024
53fd09a399c9afcd820e	b09ff1ab1d0583aa955e
ad93cb9499239edb202a	7127cdaa6d5aac38b1ab
1c95850b2ece593eeb02	6e6ca8c1794636815991
1c95850b2ece593eeb02	26cbf254ce12996ba47a
0da1a8e4012fbe1989cc	5f95f3434be5648b5d80
0da1a8e4012fbe1989cc	fd1e58160c741b9cbf3a
83db990ac81dfd8b35c7	d03263d74be35d68b7e7
74b1f4902d76a7737fad	7f501fe35061978683a9
7a6c12bdaf5c6da7b21d	158d15e754e983031efe
6128f2c7422b743e9a52	77fe10b9e66e0427adda
26f71e77ebb5d7b35f70	e9913bf2396436bfd948
8899967fd022bd010f96	0f2cb4f68319568db9ac
f11f70d30bce575096ec	5b9f00952b7726608a95
3810d251d2a035b1de85	0f2cb4f68319568db9ac
457e892a683f70a78c8a	e16f16d7233090cb7965
9a64dc0d5caa2750bf59	fd1e58160c741b9cbf3a
471952423957373168f7	513de73970f569aa0344
0b701012c2790c7c4735	7127cdaa6d5aac38b1ab
1f2141b07765892cb67a	f150e71801bd8410d574
91df72a6767bb7c3ab87	2e42bbb99dd33754220a
91df72a6767bb7c3ab87	26cbf254ce12996ba47a
7e5ffa821f50c0bc8f8a	5f95f3434be5648b5d80
862d7a9266b086d90ae6	0aaeed216fcaa328c260
117e47621ce29f1a2bc4	e5bd8960caeadda4ba86
0371cd4bdfd4b24f8cbe	20e179146eb6bd982565
877d41f9b7669c490ac5	77fe10b9e66e0427adda
97d9dad9b451c17681bc	20e179146eb6bd982565
f37759490150fb4b8c33	20e179146eb6bd982565
b8232c40c37937eb01d7	e3c5fa508ad4fe4f19d6
ed07ec8fa564381ea093	e9913bf2396436bfd948
3c98fafabb219e9ecd4d	32ee1105401237bff687
b95cadce5af0089d8b98	0f2cb4f68319568db9ac
b07f3d9674f357d4fd0d	7df17fc3ee9767f27492
faf6df81a0493135c916	a861497b23a650f9b440
67b3b7100d976861d8b7	11a6a61bf1f0c25b9e92
2bd14e1ccaaa323968fe	3ab4e5dd7f5007d04f91
3d2845b4e39b80d24680	dc3ba9a20e9d40aff0b4
b96a13884d9cb2f18ad3	5b9f00952b7726608a95
99e6504888bd1d9f9986	e16f16d7233090cb7965
c90d77745ab6d0e48537	fd1e58160c741b9cbf3a
bd6be185e5e6e11ada1b	513de73970f569aa0344
2f7f607a87c99c6722e9	7127cdaa6d5aac38b1ab
dd1d596229411ff299e5	f3a9ac717be348cb8932
9bb832cf61494f6878ed	b4b61e4d64204228f13a
9bb832cf61494f6878ed	26cbf254ce12996ba47a
04dd4b67bcf35901231b	fd1e58160c741b9cbf3a
74f2bc4a5877a6277d7c	d03263d74be35d68b7e7
649153b9a1951bb4bb9a	7f501fe35061978683a9
76d5185f429653ed2e87	20e179146eb6bd982565
bb3d38b9bf77a8e95b79	77fe10b9e66e0427adda
431824f1591821900428	7127cdaa6d5aac38b1ab
dd3c19d41c8dd7d27682	7127cdaa6d5aac38b1ab
8427cb0e609c6395011d	e16f16d7233090cb7965
13cf5f07e77de99dcc30	fd1e58160c741b9cbf3a
cf5177242c1034802d8e	513de73970f569aa0344
7f18686d985523c66aa3	7127cdaa6d5aac38b1ab
b232a97de45a99e40e65	f150e71801bd8410d574
5c960d1023e68e8b8d68	6e6ca8c1794636815991
50539e1ab2fe53f5df15	5f95f3434be5648b5d80
86c46654bedeb5a02591	0aaeed216fcaa328c260
0f692345829560ccc163	ba907005863effc8887e
ec964e12ec298f154906	158d15e754e983031efe
e48a2537426fd384e04b	77fe10b9e66e0427adda
028f24b1257fd6f8b5af	158d15e754e983031efe
0a30a4d85b028245ef62	158d15e754e983031efe
3ce1e9bfc95928b110dc	e3c5fa508ad4fe4f19d6
c5aa80e1a22c4b09fbfe	e9913bf2396436bfd948
4d780077a193e1d55df4	e16f16d7233090cb7965
809db13c3bb696a97ee7	fd1e58160c741b9cbf3a
1094e4c9fe718bcb6290	513de73970f569aa0344
ab845c01757acc0d395d	7127cdaa6d5aac38b1ab
b2e5e7431c984ac1c0d7	f3a9ac717be348cb8932
f6eeb477d5f1c37726c6	ba907005863effc8887e
f6eeb477d5f1c37726c6	26cbf254ce12996ba47a
81f20f3be795fb6d955d	5f95f3434be5648b5d80
9e72e29c3c0b838bfc4c	0aaeed216fcaa328c260
19cc8e8396329ca6c0bc	7f501fe35061978683a9
65d085a17aba3a9b1d94	158d15e754e983031efe
5d13ab8d40255bc00b5e	77fe10b9e66e0427adda
acb5f26d24423de54ac7	e3c5fa508ad4fe4f19d6
4e4f54ae37e0e224fdf3	e3c5fa508ad4fe4f19d6
\.


--
-- Data for Name: _RoleToUser; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."_RoleToUser" ("A", "B") FROM stdin;
role_admin_001	user_admin_001
role_hod_001	user_admin_001
role_teacher_001	user_admin_001
role_hod_001	v152mbwu3vdblom8o3xwoad5
role_teacher_001	v152mbwu3vdblom8o3xwoad5
role_teacher_001	7df17fc3ee9767f27492
role_teacher_001	9b6a3472cedfe4f10826
role_teacher_001	1b98a78cee88b43a5024
role_teacher_001	f150e71801bd8410d574
role_teacher_001	f1a15047dcdc49b0d1bf
role_teacher_001	b09ff1ab1d0583aa955e
role_teacher_001	5f95f3434be5648b5d80
role_teacher_001	f3a9ac717be348cb8932
role_teacher_001	513de73970f569aa0344
role_teacher_001	a861497b23a650f9b440
role_teacher_001	7f501fe35061978683a9
role_teacher_001	6298d554d8e14067ff3e
role_teacher_001	e16f16d7233090cb7965
role_teacher_001	c89abddc6d32037da55d
role_teacher_001	dc3ba9a20e9d40aff0b4
role_teacher_001	d03263d74be35d68b7e7
role_teacher_001	20e179146eb6bd982565
role_teacher_001	e3c5fa508ad4fe4f19d6
role_teacher_001	5b9f00952b7726608a95
role_teacher_001	eb100bda8b0d3ecc0c62
role_teacher_001	2e42bbb99dd33754220a
role_teacher_001	0f2cb4f68319568db9ac
role_teacher_001	6e6ca8c1794636815991
role_teacher_001	0aaeed216fcaa328c260
role_teacher_001	32ee1105401237bff687
role_teacher_001	72b3ecb6bb3ddf696ffc
role_teacher_001	3ab4e5dd7f5007d04f91
role_teacher_001	11a6a61bf1f0c25b9e92
role_teacher_001	ba907005863effc8887e
role_teacher_001	77fe10b9e66e0427adda
role_teacher_001	e5bd8960caeadda4ba86
role_teacher_001	4dcdd75becac3c5ac374
role_teacher_001	26cbf254ce12996ba47a
role_teacher_001	b4b61e4d64204228f13a
role_teacher_001	fd0e10eec0a7cb76bd4f
role_teacher_001	e9913bf2396436bfd948
role_teacher_001	158d15e754e983031efe
role_teacher_001	fd1e58160c741b9cbf3a
role_teacher_001	7127cdaa6d5aac38b1ab
role_hod_001	7127cdaa6d5aac38b1ab
\.


--
-- Data for Name: _SubjectToUser; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."_SubjectToUser" ("A", "B") FROM stdin;
\.


--
-- Data for Name: _prisma_migrations; Type: TABLE DATA; Schema: public; Owner: -
--

COPY "public"."_prisma_migrations" ("id", "checksum", "finished_at", "migration_name", "logs", "rolled_back_at", "started_at", "applied_steps_count") FROM stdin;
3a803b17-5813-420d-8665-3c31d3bcbded	200d3ab3494bc516066fd9e4ac0ac928f5d7ea5f49a2b6ae57d56cc3cdadfc5d	2026-02-02 14:10:44.506336+00	20260201013402_initial_postgresql_migration	\N	\N	2026-02-02 14:10:44.454539+00	1
f2a97987-a40e-4308-994f-0a778e8bf89c	c071a32091df8cac2d286240d05449347cbd3b8b8f8310ea5265cb024e0258d1	2026-02-02 14:10:44.513074+00	20260201013512_add_title_to_comment_group	\N	\N	2026-02-02 14:10:44.507729+00	1
c4148fe0-e8b4-4913-b0ac-c0cbada184c6	971f84ad4b96866754985a7f64779b0e470c0105dcf478bdfea263eb1aa9a174	2026-02-02 14:10:44.520249+00	20260201124425_add_form_to_pupil	\N	\N	2026-02-02 14:10:44.51466+00	1
8a30dd28-d6bc-44e8-af97-6e0d700e80bf	c23171b2ccc9d1cef583c7f727e2b6fee7804583900299d200a0661bc6b81138	2026-02-02 14:18:03.175583+00	20260202132855_add_check_status_to_assignment	\N	\N	2026-02-02 14:18:03.162505+00	1
5d943947-5507-4a73-aa34-9f3e6ad01661	4a40e0536a65b50fb00eb48484a1a78bda0d8806b4b272ab8d1e7b9a019c86da	2026-02-02 16:49:06.721234+00	20260202163708_add_deadline_model	\N	\N	2026-02-02 16:49:06.708901+00	1
d84b5cdf-ebbe-492d-adab-7c419ec90393	b9f84c2294feaa8765ec3efc3416a84a657da193edce429f067d52c46fa9193b	2026-02-03 14:33:41.989898+00	20260203004532_add_audit_log	\N	\N	2026-02-03 14:33:41.978594+00	1
0e61a92a-5754-46b8-96d2-f1cc3ade7533	fbe7dbf19171b954266fb0bb9b7738e4b55a959c5e3d1f4de4ee09bf49a4678a	2026-02-03 16:26:29.616333+00	20260203162300_add_user_is_active	\N	\N	2026-02-03 16:26:29.606939+00	1
1ba9c5cc-13c3-4212-bda1-eb799c3f884f	79d1d7094b42d37b10ba78faa9a517342383b2e978821ce888fc0369677ef05a	2026-02-14 14:48:47.62045+00	20260206011712_add_common_comment_groups	\N	\N	2026-02-14 14:48:47.591104+00	1
585a6ffc-4ce2-433b-a2a2-cb670faa8c77	ab22cf37b5db1e5cdc28d093c969e16a8367586ca95a28daa74fc7071eb3c3e9	2026-02-14 14:48:47.629712+00	20260206081505_add_linked_comment_groups	\N	\N	2026-02-14 14:48:47.621783+00	1
b69e7db9-c51f-4323-90c6-800495445ac9	5e20f48169d72b9ec7556ba4d6c905e6b06e18117bfeb272a3a6db09be365410	2026-02-14 14:48:47.638144+00	20260214132249_remove_paragraph_position_add_comment_format	\N	\N	2026-02-14 14:48:47.631146+00	1
5843f172-ceac-4fb3-8cc3-b5579ce1a212	9a9a56d24677cdfb6d2cadcbd0a2704f7098b770a4c4c6e522fc77f00bb260d0	2026-03-02 02:59:46.707271+00	20260301000000_add_comment_format_to_subject	\N	\N	2026-03-02 02:59:46.483994+00	1
\.


--
-- Name: AppSetting AppSetting_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."AppSetting"
    ADD CONSTRAINT "AppSetting_pkey" PRIMARY KEY ("key");


--
-- Name: Assignment Assignment_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."Assignment"
    ADD CONSTRAINT "Assignment_pkey" PRIMARY KEY ("id");


--
-- Name: AuditLog AuditLog_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."AuditLog"
    ADD CONSTRAINT "AuditLog_pkey" PRIMARY KEY ("id");


--
-- Name: Class Class_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."Class"
    ADD CONSTRAINT "Class_pkey" PRIMARY KEY ("id");


--
-- Name: CommentGroup CommentGroup_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."CommentGroup"
    ADD CONSTRAINT "CommentGroup_pkey" PRIMARY KEY ("id");


--
-- Name: CommentOption CommentOption_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."CommentOption"
    ADD CONSTRAINT "CommentOption_pkey" PRIMARY KEY ("id");


--
-- Name: CommonCommentGroup CommonCommentGroup_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."CommonCommentGroup"
    ADD CONSTRAINT "CommonCommentGroup_pkey" PRIMARY KEY ("id");


--
-- Name: CommonCommentOption CommonCommentOption_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."CommonCommentOption"
    ADD CONSTRAINT "CommonCommentOption_pkey" PRIMARY KEY ("id");


--
-- Name: CommonPupilCode CommonPupilCode_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."CommonPupilCode"
    ADD CONSTRAINT "CommonPupilCode_pkey" PRIMARY KEY ("id");


--
-- Name: Deadline Deadline_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."Deadline"
    ADD CONSTRAINT "Deadline_pkey" PRIMARY KEY ("id");


--
-- Name: IgnoredWord IgnoredWord_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."IgnoredWord"
    ADD CONSTRAINT "IgnoredWord_pkey" PRIMARY KEY ("id");


--
-- Name: IgnoredWord IgnoredWord_teacherId_word_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."IgnoredWord"
    ADD CONSTRAINT "IgnoredWord_teacherId_word_key" UNIQUE ("teacherId", "word");


--
-- Name: PupilCode PupilCode_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."PupilCode"
    ADD CONSTRAINT "PupilCode_pkey" PRIMARY KEY ("id");


--
-- Name: Pupil Pupil_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."Pupil"
    ADD CONSTRAINT "Pupil_pkey" PRIMARY KEY ("admissionNumber");


--
-- Name: Role Role_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."Role"
    ADD CONSTRAINT "Role_pkey" PRIMARY KEY ("id");


--
-- Name: Subject Subject_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."Subject"
    ADD CONSTRAINT "Subject_pkey" PRIMARY KEY ("id");


--
-- Name: User User_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."User"
    ADD CONSTRAINT "User_pkey" PRIMARY KEY ("id");


--
-- Name: _ClassToUser _ClassToUser_AB_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."_ClassToUser"
    ADD CONSTRAINT "_ClassToUser_AB_pkey" PRIMARY KEY ("A", "B");


--
-- Name: _RoleToUser _RoleToUser_AB_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."_RoleToUser"
    ADD CONSTRAINT "_RoleToUser_AB_pkey" PRIMARY KEY ("A", "B");


--
-- Name: _SubjectToUser _SubjectToUser_AB_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."_SubjectToUser"
    ADD CONSTRAINT "_SubjectToUser_AB_pkey" PRIMARY KEY ("A", "B");


--
-- Name: _prisma_migrations _prisma_migrations_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."_prisma_migrations"
    ADD CONSTRAINT "_prisma_migrations_pkey" PRIMARY KEY ("id");


--
-- Name: Assignment_checkStatus_idx; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX "Assignment_checkStatus_idx" ON "public"."Assignment" USING "btree" ("checkStatus");


--
-- Name: Assignment_classId_idx; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX "Assignment_classId_idx" ON "public"."Assignment" USING "btree" ("classId");


--
-- Name: Assignment_pupilId_idx; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX "Assignment_pupilId_idx" ON "public"."Assignment" USING "btree" ("pupilId");


--
-- Name: AuditLog_action_idx; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX "AuditLog_action_idx" ON "public"."AuditLog" USING "btree" ("action");


--
-- Name: AuditLog_createdAt_idx; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX "AuditLog_createdAt_idx" ON "public"."AuditLog" USING "btree" ("createdAt");


--
-- Name: AuditLog_entityType_idx; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX "AuditLog_entityType_idx" ON "public"."AuditLog" USING "btree" ("entityType");


--
-- Name: AuditLog_userId_idx; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX "AuditLog_userId_idx" ON "public"."AuditLog" USING "btree" ("userId");


--
-- Name: Class_name_key; Type: INDEX; Schema: public; Owner: -
--

CREATE UNIQUE INDEX "Class_name_key" ON "public"."Class" USING "btree" ("name");


--
-- Name: Class_subjectId_idx; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX "Class_subjectId_idx" ON "public"."Class" USING "btree" ("subjectId");


--
-- Name: CommentGroup_subjectId_name_key; Type: INDEX; Schema: public; Owner: -
--

CREATE UNIQUE INDEX "CommentGroup_subjectId_name_key" ON "public"."CommentGroup" USING "btree" ("subjectId", "name");


--
-- Name: CommentOption_groupId_code_key; Type: INDEX; Schema: public; Owner: -
--

CREATE UNIQUE INDEX "CommentOption_groupId_code_key" ON "public"."CommentOption" USING "btree" ("groupId", "code");


--
-- Name: CommentOption_groupId_idx; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX "CommentOption_groupId_idx" ON "public"."CommentOption" USING "btree" ("groupId");


--
-- Name: CommonCommentGroup_name_key; Type: INDEX; Schema: public; Owner: -
--

CREATE UNIQUE INDEX "CommonCommentGroup_name_key" ON "public"."CommonCommentGroup" USING "btree" ("name");


--
-- Name: CommonCommentOption_groupId_code_key; Type: INDEX; Schema: public; Owner: -
--

CREATE UNIQUE INDEX "CommonCommentOption_groupId_code_key" ON "public"."CommonCommentOption" USING "btree" ("groupId", "code");


--
-- Name: CommonCommentOption_groupId_idx; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX "CommonCommentOption_groupId_idx" ON "public"."CommonCommentOption" USING "btree" ("groupId");


--
-- Name: CommonPupilCode_assignmentId_commonGroupId_key; Type: INDEX; Schema: public; Owner: -
--

CREATE UNIQUE INDEX "CommonPupilCode_assignmentId_commonGroupId_key" ON "public"."CommonPupilCode" USING "btree" ("assignmentId", "commonGroupId");


--
-- Name: CommonPupilCode_assignmentId_idx; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX "CommonPupilCode_assignmentId_idx" ON "public"."CommonPupilCode" USING "btree" ("assignmentId");


--
-- Name: CommonPupilCode_commonGroupId_idx; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX "CommonPupilCode_commonGroupId_idx" ON "public"."CommonPupilCode" USING "btree" ("commonGroupId");


--
-- Name: IgnoredWord_teacherId_idx; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX "IgnoredWord_teacherId_idx" ON "public"."IgnoredWord" USING "btree" ("teacherId");


--
-- Name: PupilCode_assignmentId_groupId_key; Type: INDEX; Schema: public; Owner: -
--

CREATE UNIQUE INDEX "PupilCode_assignmentId_groupId_key" ON "public"."PupilCode" USING "btree" ("assignmentId", "groupId");


--
-- Name: PupilCode_assignmentId_idx; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX "PupilCode_assignmentId_idx" ON "public"."PupilCode" USING "btree" ("assignmentId");


--
-- Name: PupilCode_groupId_idx; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX "PupilCode_groupId_idx" ON "public"."PupilCode" USING "btree" ("groupId");


--
-- Name: Role_name_key; Type: INDEX; Schema: public; Owner: -
--

CREATE UNIQUE INDEX "Role_name_key" ON "public"."Role" USING "btree" ("name");


--
-- Name: Subject_code_key; Type: INDEX; Schema: public; Owner: -
--

CREATE UNIQUE INDEX "Subject_code_key" ON "public"."Subject" USING "btree" ("code");


--
-- Name: User_username_key; Type: INDEX; Schema: public; Owner: -
--

CREATE UNIQUE INDEX "User_username_key" ON "public"."User" USING "btree" ("username");


--
-- Name: _ClassToUser_B_index; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX "_ClassToUser_B_index" ON "public"."_ClassToUser" USING "btree" ("B");


--
-- Name: _RoleToUser_B_index; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX "_RoleToUser_B_index" ON "public"."_RoleToUser" USING "btree" ("B");


--
-- Name: _SubjectToUser_B_index; Type: INDEX; Schema: public; Owner: -
--

CREATE INDEX "_SubjectToUser_B_index" ON "public"."_SubjectToUser" USING "btree" ("B");


--
-- Name: Assignment Assignment_checkedById_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."Assignment"
    ADD CONSTRAINT "Assignment_checkedById_fkey" FOREIGN KEY ("checkedById") REFERENCES "public"."User"("id") ON UPDATE CASCADE ON DELETE SET NULL;


--
-- Name: Assignment Assignment_classId_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."Assignment"
    ADD CONSTRAINT "Assignment_classId_fkey" FOREIGN KEY ("classId") REFERENCES "public"."Class"("id") ON UPDATE CASCADE ON DELETE CASCADE;


--
-- Name: Assignment Assignment_pupilId_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."Assignment"
    ADD CONSTRAINT "Assignment_pupilId_fkey" FOREIGN KEY ("pupilId") REFERENCES "public"."Pupil"("admissionNumber") ON UPDATE CASCADE ON DELETE CASCADE;


--
-- Name: Class Class_subjectId_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."Class"
    ADD CONSTRAINT "Class_subjectId_fkey" FOREIGN KEY ("subjectId") REFERENCES "public"."Subject"("id") ON UPDATE CASCADE ON DELETE CASCADE;


--
-- Name: CommentGroup CommentGroup_subjectId_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."CommentGroup"
    ADD CONSTRAINT "CommentGroup_subjectId_fkey" FOREIGN KEY ("subjectId") REFERENCES "public"."Subject"("id") ON UPDATE CASCADE ON DELETE CASCADE;


--
-- Name: CommentOption CommentOption_groupId_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."CommentOption"
    ADD CONSTRAINT "CommentOption_groupId_fkey" FOREIGN KEY ("groupId") REFERENCES "public"."CommentGroup"("id") ON UPDATE CASCADE ON DELETE CASCADE;


--
-- Name: CommonCommentOption CommonCommentOption_groupId_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."CommonCommentOption"
    ADD CONSTRAINT "CommonCommentOption_groupId_fkey" FOREIGN KEY ("groupId") REFERENCES "public"."CommonCommentGroup"("id") ON UPDATE CASCADE ON DELETE CASCADE;


--
-- Name: CommonPupilCode CommonPupilCode_assignmentId_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."CommonPupilCode"
    ADD CONSTRAINT "CommonPupilCode_assignmentId_fkey" FOREIGN KEY ("assignmentId") REFERENCES "public"."Assignment"("id") ON UPDATE CASCADE ON DELETE CASCADE;


--
-- Name: CommonPupilCode CommonPupilCode_commonGroupId_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."CommonPupilCode"
    ADD CONSTRAINT "CommonPupilCode_commonGroupId_fkey" FOREIGN KEY ("commonGroupId") REFERENCES "public"."CommonCommentGroup"("id") ON UPDATE CASCADE ON DELETE CASCADE;


--
-- Name: IgnoredWord IgnoredWord_teacherId_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."IgnoredWord"
    ADD CONSTRAINT "IgnoredWord_teacherId_fkey" FOREIGN KEY ("teacherId") REFERENCES "public"."User"("id") ON DELETE CASCADE;


--
-- Name: PupilCode PupilCode_assignmentId_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."PupilCode"
    ADD CONSTRAINT "PupilCode_assignmentId_fkey" FOREIGN KEY ("assignmentId") REFERENCES "public"."Assignment"("id") ON UPDATE CASCADE ON DELETE CASCADE;


--
-- Name: PupilCode PupilCode_groupId_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."PupilCode"
    ADD CONSTRAINT "PupilCode_groupId_fkey" FOREIGN KEY ("groupId") REFERENCES "public"."CommentGroup"("id") ON UPDATE CASCADE ON DELETE CASCADE;


--
-- Name: _ClassToUser _ClassToUser_A_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."_ClassToUser"
    ADD CONSTRAINT "_ClassToUser_A_fkey" FOREIGN KEY ("A") REFERENCES "public"."Class"("id") ON UPDATE CASCADE ON DELETE CASCADE;


--
-- Name: _ClassToUser _ClassToUser_B_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."_ClassToUser"
    ADD CONSTRAINT "_ClassToUser_B_fkey" FOREIGN KEY ("B") REFERENCES "public"."User"("id") ON UPDATE CASCADE ON DELETE CASCADE;


--
-- Name: _RoleToUser _RoleToUser_A_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."_RoleToUser"
    ADD CONSTRAINT "_RoleToUser_A_fkey" FOREIGN KEY ("A") REFERENCES "public"."Role"("id") ON UPDATE CASCADE ON DELETE CASCADE;


--
-- Name: _RoleToUser _RoleToUser_B_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."_RoleToUser"
    ADD CONSTRAINT "_RoleToUser_B_fkey" FOREIGN KEY ("B") REFERENCES "public"."User"("id") ON UPDATE CASCADE ON DELETE CASCADE;


--
-- Name: _SubjectToUser _SubjectToUser_A_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."_SubjectToUser"
    ADD CONSTRAINT "_SubjectToUser_A_fkey" FOREIGN KEY ("A") REFERENCES "public"."Subject"("id") ON UPDATE CASCADE ON DELETE CASCADE;


--
-- Name: _SubjectToUser _SubjectToUser_B_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY "public"."_SubjectToUser"
    ADD CONSTRAINT "_SubjectToUser_B_fkey" FOREIGN KEY ("B") REFERENCES "public"."User"("id") ON UPDATE CASCADE ON DELETE CASCADE;


--
COMMIT;

-- PostgreSQL database dump complete
--


