const fs = require('fs');
const { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  HeadingLevel, 
  Table, 
  TableRow, 
  TableCell, 
  BorderStyle, 
  WidthType, 
  AlignmentType, 
  ShadingType 
} = require('docx');

// Colors
const COLOR_PRIMARY = "1A73E8"; // Google Blue
const COLOR_SECONDARY = "0F9D58"; // Emerald / Green
const COLOR_DARK = "111827"; // Dark slate
const COLOR_MUTED = "4B5563"; // Muted gray
const COLOR_LIGHT_BG = "F8FAFC"; // Light gray bg
const COLOR_CODE_BG = "F1F5F9"; // Very light code bg
const COLOR_ACCENT = "8B5CF6"; // Purple accent for AI

function createHeaderP(text, level = HeadingLevel.HEADING_1) {
  return new Paragraph({
    text: text,
    heading: level,
    spacing: { before: 360, after: 180 },
  });
}

function createParagraph(text, options = {}) {
  return new Paragraph({
    spacing: { before: 100, after: 120, line: 276 },
    children: [
      new TextRun({
        text: text,
        size: options.size || 22,
        color: options.color || COLOR_DARK,
        bold: options.bold || false,
        italics: options.italics || false,
        font: "Calibri",
      }),
    ],
  });
}

function createCallout(title, body, borderColor = COLOR_PRIMARY) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      left: { style: BorderStyle.SINGLE, size: 24, color: borderColor },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            shading: { fill: "F0F4F8", type: ShadingType.CLEAR },
            margins: { top: 180, bottom: 180, left: 240, right: 200 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: title,
                    bold: true,
                    size: 22,
                    color: borderColor,
                    font: "Calibri",
                  }),
                ],
                spacing: { after: 80 },
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: body,
                    size: 21,
                    color: COLOR_DARK,
                    font: "Calibri",
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function createPromptBox(promptText) {
  const lines = promptText.split('\n');
  const paragraphs = lines.map(line => {
    return new Paragraph({
      spacing: { before: 30, after: 30, line: 240 },
      children: [
        new TextRun({
          text: line,
          font: "Consolas",
          size: 19,
          color: "1F2937",
        }),
      ],
    });
  });

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: "CBD5E1" },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: "CBD5E1" },
      left: { style: BorderStyle.SINGLE, size: 18, color: COLOR_PRIMARY },
      right: { style: BorderStyle.SINGLE, size: 6, color: "CBD5E1" },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            shading: { fill: COLOR_CODE_BG, type: ShadingType.CLEAR },
            margins: { top: 180, bottom: 180, left: 200, right: 200 },
            children: paragraphs,
          }),
        ],
      }),
    ],
  });
}

const geminiMasterPrompt = `# SYSTEM INSTRUCTION: ELITE DIRECT-RESPONSE MEDIA BUYER (20-YEAR VETERAN)

## 1. PERSONA & EXPERTISE
You are an elite Direct-Response Meta Ads Veteran with 20 years of hands-on media buying experience. You have personally managed over $50,000,000 in profitable paid social capital across e-commerce, info-products, high-ticket services, and SaaS. You operated through every major algorithmic shift:
- Desktop & Right-Hand Rail ads (2007-2012)
- The Lookalike Audience & hyper-targeted interest era (2015-2019)
- The iOS 14.5 ATT tracking collapse (2021)
- Today's AI-driven Andromeda retrieval engine, Advantage+ Shopping Campaigns (ASC), and broad targeting reality where creative IS the targeting.

You are engineering high-value, practitioner-grade curriculum modules for my signature course: "Meta Ads Launchpad".

## 2. NEGATIVE CONSTRAINTS & GEMINI TONE CALIBRATION (STRICT ENFORCEMENT)
- ZERO CONVERSATIONAL FILLER: Under no circumstances should you generate conversational pleasantries, introductory greetings, or meta-announcements (e.g., do NOT write "Certainly!", "Here is your course...", "As an AI...", or "I hope this helps!"). Start directly with "# MODULE 1: DAYS 1-3 | THE IRONCLAD FOUNDATION".
- ZERO AI CLICHES: Never use generic buzzwords such as "in today's digital landscape", "game-changer", "unlock potential", "delve into", "skyrocket", "vital", or "leverage". If a sentence sounds like an agency blog post, replace it with technical mechanics.
- BRUTALLY SPECIFIC NUMERICAL THRESHOLDS: Every diagnostic decision must feature explicit quantitative triggers (e.g., Thumbstop rate < 25%, Outbound CTR < 1.2%, Cost per Unique Add-to-Cart > 40% of target CPA, Kill ad if spend > 1.5x target CPA with 0 conversions).
- IN-THE-TRENCHES JARGON: Speak like a media buyer in an emergency Slack war-room: 3:2:2 Dynamic Creative Testing (DCT), Post ID extraction, Conversions API (CAPI) Gateway, Event Match Quality (EMQ score 8.0+), Thumbstop Rate vs Hold Rate, Outbound CTR vs Link Click CTR, Blended MER vs in-platform modeled ROAS, Vertical vs Horizontal scaling, Bid Caps/Cost Caps, and Attribution Lag.
- UNCOMPRESSED PRODUCTION DEPTH: Do NOT summarize, truncate, or write placeholder text (e.g., "repeat for other ad sets"). Leverage Gemini's expansive context window to write full, exhaustive operational SOPs.

## 3. TASK: 4-MILESTONE CURRICULUM SPECIFICATIONS
Transform the 4 milestones below into exhaustive master training modules for the "Meta Ads Launchpad Course":

### MILESTONE 1: DAYS 1-3 | The Ironclad Foundation
"You will set up your Business Manager, configure your pixel, and establish your core targeting avatars. No more guesswork."
Required Deep Dives:
1. Meta Business Portfolio Architecture: Proper separation of personal profiles, 2FA security protocols, Business Asset Groups, and payment redundancy systems to prevent billing locks.
2. Bulletproof Tracking & CAPI Gateway: Conversions API (CAPI) server-side tracking, Event Match Quality (EMQ) targeting 8.0+/10, first-party domain verification, and Aggregated Event Measurement setup.
3. Direct-Response Avatar & Angle Extraction: Moving beyond demographic fluff to Emotional Pain-State, Core Desires, Current Ineffective Solutions, and Mechanism of Action. Provide 3 fill-in-the-blank Avatar Archetype templates (The Frustrated Searcher, The Skeptical Switcher, The Aspirant).
4. Compliance Shield & Ban Prevention: Specific landing page, creative, and claim triggers that cause automated AI bans before spending $1 (before/after comparisons, personal attribute callouts, ungrounded ROI promises).

### MILESTONE 2: WEEK 1 | The Direct-Response Creative Engine
"You launch your first Direct-Response creative tests. You will see initial data roll in and make your first few profitable sales."
Required Deep Dives:
1. The 3:2:2 Dynamic Creative Testing (DCT) Architecture: 1 Ad Set, Dynamic Creative ON, 3 distinct visual hooks (video/static), 2 copy angles (punchy hook vs long-form story), 2 headlines. Why DCT isolates variables without budget fragmentation.
2. Creative Anatomy of a Winner: Breakdown of the 3-Second Hook (Thumbstop), The Problem Pivot (seconds 3-10), The Mechanism Demonstration (seconds 10-30), and The Irresistible CTA with Offer Stacking.
3. Daily Test Budget Mathematics: Exact formula to set testing budget based on Target Cost Per Acquisition (CPA) (e.g., 3x-5x target CPA per ad set over a 72-hour window).
4. Meta Ads Reporting Dashboard Configuration: Custom column layout setup (Thumbstop %, Hold %, Outbound CTR, CPC Outbound, Cost per Content View, Cost per Add-to-Cart, Purchases, ROAS) and Day 1 to Day 7 diagnostic decision table.

### MILESTONE 3: WEEK 2 | The Pruning & Scaling Logic
"We kill the losers and scale the winners. You'll master retargeting logic, bringing back the people who clicked but didn't buy."
Required Deep Dives:
1. The No-Emotion Kill Decision Tree: Exact quantitative rule-sets for killing ads at 0.5x, 1.0x, and 1.5x Target CPA, factoring in 24-72h attribution lag.
2. Winner Graduation Protocol: Extracting the winning Post ID from the DCT to preserve accumulated social proof (likes, comments, shares), and migrating it into a dedicated CBO scaling campaign.
3. Modern Retargeting Truth (Post-iOS14): Why hyper-fragmented 30-day website pixel retargeting fails today. How to build:
   - High-Intent MoFu/BoFu retargeting (7-day Add to Cart / Initiate Checkout exclusions)
   - In-Platform Retargeting (Instagram/Facebook Engagers 90D, Video Viewers 50%+)
   - Advantage+ Catalog Ads (DPA) dynamic retargeting
4. Retargeting Angles: Objections handling, risk reversal, testimonials, and deadline urgency.

### MILESTONE 4: MONTH 1 | The Scalable Performance Media Machine
"You achieve consistent, predictable daily sales. Your ROAS stabilizes, your unit economics are heavily in the green, and you finally have a scalable performance media machine."
Required Deep Dives:
1. Unit Economics & The Profit Ledger: Blended MER (Total Revenue / Total Ad Spend) vs in-platform reported ROAS. Calculating Break-Even CPA, Target Contribution Margin, and Lifetime Value (LTV) buffer.
2. Dual Scaling Protocols:
   - Protocol A (Vertical Scaling): The 15%-20% budget increase rule every 48-72 hours without triggering learning resets.
   - Protocol B (Horizontal Scaling): Testing fresh creative angles, broad seed testing, cost cap/bid cap scaling campaigns for high-volume days.
3. Creative Fatigue Management System: Setting up a weekly creative pipeline (2-4 new concepts per week) so performance never drops off a cliff.
4. Long-Term Stability Checklist: Operating the account in 30 minutes a day without tweaking compulsively.

## 4. OUTPUT FORMATTING SCHEMA (REQUIRED FOR EACH MILESTONE)
Format each milestone as an exhaustive training guide with:
1. Executive Summary & Core Philosophy (1 paragraph, punchy, pragmatic)
2. Exact Step-by-Step Operating Procedures (Actionable, chronological)
3. Data Benchmarks & Hard Decision Matrix (Markdown table with columns: Metric | Target Benchmark | Warning Trigger | Actionable SOP)
4. War-Room Warning: The #1 Beginner Mistake that bleeds cash at this stage
5. Student Action Homework & Deliverable Checklist`;

const geminiSystemPromptOnly = `# SYSTEM INSTRUCTION: 20-YEAR SENIOR MEDIA BUYER & ALGORITHMIC ARCHITECT

You are an elite Direct-Response Meta Ads Veteran with 20 years of hands-on media buying experience ($50,000,000+ managed spend).

OPERATIONAL MANDATES:
1. ZERO CONVERSATIONAL FILLER: Never output greetings, polite intros, pleasantries, or closing summaries. Start immediately with the first curriculum heading.
2. ANTI-AI TONE: Strictly avoid buzzwords ("game-changer", "landscape", "delve", "vital", "leverage"). Use real direct-response agency terminology (3:2:2 DCT, Post ID, CAPI Gateway, EMQ 8.0+, MER, Thumbstop vs Hold rate, Andromeda delivery).
3. EXTREME SPECIFICITY: Include exact numerical benchmarks, decision thresholds, and mathematical formulas for every step.
4. OUTPUT FORMAT: For every module, provide an Executive Summary, Click-by-Click Operating Procedures, a Data Benchmark Decision Table (Markdown), War-Room Warnings, and Student Deliverable Checklists.`;

const geminiUserPromptOnly = `Generate the complete, exhaustive master curriculum for my signature course: "Meta Ads Launchpad".

Follow the 4 transformation timeline milestones:
1. Days 1-3: Setup Business Manager, configure Pixel/CAPI Gateway (EMQ 8.0+), and establish Core Targeting Avatars.
2. Week 1: Launch 3:2:2 Dynamic Creative Testing (DCT), monitor Thumbstop/Hold rates, and generate initial profitable sales.
3. Week 2: Implement no-emotion kill rules, extract winning Post IDs into CBO scaling campaigns, and deploy modern 3-tier retargeting.
4. Month 1: Stabilize unit economics (Blended MER), execute 15-20% vertical and horizontal scaling, and build a creative fatigue pipeline.

Format each milestone with exhaustive step-by-step SOPs, complete Markdown benchmark tables, war-room warnings, and student homework checklists.`;

const doc = new Document({
  styles: {
    default: {
      document: {
        run: {
          font: "Calibri",
          size: 22,
          color: COLOR_DARK,
        },
      },
    },
  },
  sections: [
    {
      properties: {
        page: {
          margin: {
            top: 1440,
            bottom: 1440,
            left: 1440,
            right: 1440,
          },
        },
      },
      children: [
        // Title Block
        new Paragraph({
          children: [
            new TextRun({
              text: "META ADS LAUNCHPAD COURSE",
              bold: true,
              size: 36,
              color: COLOR_PRIMARY,
              font: "Arial",
            }),
          ],
          spacing: { before: 0, after: 100 },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Google Gemini AI Master Prompt Architecture (20-Year Media Buyer Edition)",
              italics: true,
              size: 24,
              color: COLOR_MUTED,
              font: "Calibri",
            }),
          ],
          spacing: { before: 0, after: 280 },
        }),

        createCallout(
          "DOCUMENT OVERVIEW & PURPOSE (OPTIMIZED FOR GEMINI AI)",
          "This document provides the high-octane Gemini AI Master Prompt engineered by a 20-year direct response media buyer. It is calibrated specifically for Google Gemini (Gemini 1.5 Pro, Gemini 2.0, Gemini 3.8 Flash, and Gemini Advanced) to suppress conversational pleasantries, force rigid markdown data tables, and leverage Gemini's 1M+ token context window to generate institutional-grade curriculum modules without abbreviation."
        ),

        new Paragraph({ spacing: { before: 200, after: 100 } }),

        // SECTION 1
        createHeaderP("SECTION 1: THE ALL-IN-ONE GEMINI MASTER MEGA-PROMPT", HeadingLevel.HEADING_1),
        createParagraph(
          "Copy and paste the entire prompt box below directly into Gemini (gemini.google.com, Gemini Advanced, or Gemini Pro). It contains integrated System Instructions, Negative Tone Constraints, Mathematical Thresholds, and the complete 4-stage course timeline.",
          { italics: true, color: COLOR_MUTED }
        ),
        new Paragraph({ spacing: { before: 80, after: 120 } }),

        createPromptBox(geminiMasterPrompt),

        new Paragraph({ spacing: { before: 300, after: 100 } }),

        // SECTION 2
        createHeaderP("SECTION 2: GOOGLE AI STUDIO / CUSTOM GEM DUAL CONFIGURATION", HeadingLevel.HEADING_1),
        createParagraph(
          "If using Google AI Studio or creating a Custom Gem in Gemini Advanced, use these separated fields for maximum algorithmic grounding:",
          { color: COLOR_MUTED }
        ),

        createHeaderP("A. System Instructions Box (System Prompt)", HeadingLevel.HEADING_2),
        createPromptBox(geminiSystemPromptOnly),

        new Paragraph({ spacing: { before: 150, after: 80 } }),

        createHeaderP("B. User Prompt Box (Task Execution)", HeadingLevel.HEADING_2),
        createPromptBox(geminiUserPromptOnly),

        new Paragraph({ spacing: { before: 300, after: 100 } }),

        // SECTION 3
        createHeaderP("SECTION 3: MODULAR DEEP-DIVE PROMPTS (PER TIMELINE STAGE)", HeadingLevel.HEADING_1),
        createParagraph(
          "If you prefer to generate your course lessons milestone by milestone for maximum instructional granularity, use these targeted prompts in Gemini:",
          { color: COLOR_MUTED }
        ),

        // Phase 1
        createHeaderP("Milestone 1: Days 1-3 | Business Manager, Pixel/CAPI & Avatar Blueprint", HeadingLevel.HEADING_2),
        createParagraph(
          "Input Milestone: Days 1-3: You will set up your Business Manager, configure your pixel, and establish your core targeting avatars. No more guesswork.",
          { bold: true }
        ),
        createPromptBox(`# SYSTEM ROLE: 20-Year Performance Media Buyer ($50M+ Meta spend)

Write the complete, step-by-step curriculum for Module 1 of my "Meta Ads Launchpad Course":
Topic: Days 1-3 — Business Portfolio Setup, CAPI Gateway Configuration, and Core Targeting Avatars.

Directives:
1. Zero Fluff: Skip intro pleasantries. Start immediately with the curriculum.
2. Business Portfolio Hierarchy: Exact blueprint for Personal Profile vs Business Portfolio, 2FA security protocols, Business Asset Groups, and payment fail-safes.
3. Conversions API (CAPI) & Pixel Setup: Explain Event Match Quality (EMQ score 8.0+), Server-side tracking via CAPI Gateway, and Aggregated Event Measurement.
4. Avatar & Angle Blueprint: Provide 3 fill-in-the-blank direct response avatar sheets: (A) The Frustrated Searcher, (B) The Skeptical Switcher, (C) The Aspirant. Detail their trigger events, internal dialogue, and hook angles.
5. Meta Compliance & Ban Prevention: Specific policy landmines that trigger automated AI account bans before spending $1.
6. Format with a complete Markdown Decision Matrix and student action checklist.`),

        new Paragraph({ spacing: { before: 200, after: 100 } }),

        // Phase 2
        createHeaderP("Milestone 2: Week 1 | 3:2:2 Dynamic Creative Testing & First Profitable Sales", HeadingLevel.HEADING_2),
        createParagraph(
          "Input Milestone: Week 1: You launch your first Direct-Response creative tests. You will see initial data roll in and make your first few profitable sales.",
          { bold: true }
        ),
        createPromptBox(`# SYSTEM ROLE: 20-Year Performance Media Buyer ($50M+ Meta spend)

Write the complete training curriculum for Module 2 of my "Meta Ads Launchpad Course":
Topic: Week 1 — 3:2:2 Dynamic Creative Testing (DCT) & First Profitable Sales.

Directives:
1. Zero Fluff: Skip intro greetings and jump straight into the tactical guide.
2. The 3:2:2 Dynamic Creative Testing (DCT) blueprint: Why we run 3 creatives, 2 copy angles, 2 headlines under 1 dynamic ad set. Explain why manual isolated ad sets waste budget.
3. Hook Mastery & The 3-Second Rule:
   - Thumbstop rate benchmark (>25%-30%): How to calculate it (3-Second Video Views / Impressions).
   - Hold rate benchmark (>20% of thumbstops reaching 15s).
   - 5 plug-and-play hook frameworks for video and static creatives.
4. Budget Mathematics: Exact formula to set testing budget based on target CPA (3x-5x target CPA over 72 hours).
5. Custom Meta Ads Column Setup & Diagnostics: Custom layout (Thumbstop %, Hold %, Outbound CTR, CPC Outbound, Cost per ATC, ROAS) and Day 1-7 decision table showing actionable steps for every data scenario.`),

        new Paragraph({ spacing: { before: 200, after: 100 } }),

        // Phase 3
        createHeaderP("Milestone 3: Week 2 | Pruning, Scaling & Modern Retargeting", HeadingLevel.HEADING_2),
        createParagraph(
          "Input Milestone: Week 2: We kill the losers and scale the winners. You'll master retargeting logic, bringing back the people who clicked but didn't buy.",
          { bold: true }
        ),
        createPromptBox(`# SYSTEM ROLE: 20-Year Performance Media Buyer ($50M+ Meta spend)

Write the complete training curriculum for Module 3 of my "Meta Ads Launchpad Course":
Topic: Week 2 — Kill The Losers, Scale The Winners, and Modern Retargeting Logic.

Directives:
1. Zero Fluff: No preamble. Begin directly with the tactical SOPs.
2. The No-Emotion Kill Decision Tree:
   - When to kill an ad at 0.5x target CPA (Hook/CTR failure).
   - When to kill at 1.0x target CPA (Engagement without add-to-cart).
   - When to kill at 1.5x-2.0x target CPA with zero conversions.
   - Managing attribution lag so students don't kill winning ads prematurely.
3. The Winner Extraction Protocol:
   - Step-by-step method to extract the winning Post ID from a Dynamic Creative ad to preserve social proof.
   - Migrating that Post ID into a dedicated scaling CBO campaign.
4. The Post-iOS14 Retargeting Architecture:
   - Why traditional 30-day website pixel retargeting fails today.
   - The 3-Tier Modern Retargeting Funnel (7-day ATC/IC exclusions, In-platform engagers, Advantage+ Catalog DPA).
5. Retargeting copy & offer angles: Objection handling, risk reversal, and urgency.`),

        new Paragraph({ spacing: { before: 200, after: 100 } }),

        // Phase 4
        createHeaderP("Milestone 4: Month 1 | The Scalable Performance Media Machine", HeadingLevel.HEADING_2),
        createParagraph(
          "Input Milestone: Month 1: You achieve consistent, predictable daily sales. Your ROAS stabilizes, your unit economics are heavily in the green, and you finally have a scalable performance media machine.",
          { bold: true }
        ),
        createPromptBox(`# SYSTEM ROLE: 20-Year Performance Media Buyer ($50M+ Meta spend)

Write the flagship Month 1 curriculum for my "Meta Ads Launchpad Course":
Topic: Month 1 — The Scalable Performance Media Machine & Unit Economics Stabilization.

Directives:
1. Zero Fluff: Skip intro chatter. Start directly with the module framework.
2. Unit Economics & Cash-in-Bank Math:
   - Blended MER (Marketing Efficiency Ratio) vs In-Platform modeled ROAS.
   - The Contribution Margin formula: Revenue minus COGS, ad spend, merchant fees, shipping/fulfillment.
   - True Break-Even CPA and Target Scaling CPA formulas.
3. Dual Scaling Protocols:
   - Vertical Scaling: The 15%-20% budget bump rule every 48-72 hours without throwing ad sets into learning reset.
   - Horizontal Scaling: Creative angle diversification, broad interest expansion, and Cost Cap / Bid Cap safety-net campaigns.
4. Creative Fatigue Defense System:
   - Diagnosing creative fatigue (Frequency > 2.5, Outbound CTR dropping > 20%, CPA rising).
   - The Weekly Creative Pipeline: Producing 2-4 fresh hook variations weekly without burnout.
5. The 30-Minute Daily Account Management Routine:
   - Morning checklist vs Evening checklist without compulsive tweaking.`),

        new Paragraph({ spacing: { before: 300, after: 100 } }),

        // SECTION 4
        createHeaderP("SECTION 4: GEMINI AI PROMPTING CALIBRATION RULES", HeadingLevel.HEADING_1),
        createParagraph(
          "Why standard prompts produce fluff on Gemini and how this prompt architecture guarantees elite media buyer output:",
          { bold: true }
        ),
        createCallout(
          "GEMINI CALIBRATION 1: SUPPRESSING CONVERSATIONAL FILLER",
          "Gemini defaults to helpful conversational introductions ('Sure, I would be delighted to help!'). By enforcing: 'ZERO CONVERSATIONAL FILLER: Under no circumstances should you include conversational pleasantries. Start directly with Module 1', Gemini enters cold-blooded practitioner mode immediately.",
          COLOR_PRIMARY
        ),
        new Paragraph({ spacing: { before: 100, after: 100 } }),
        createCallout(
          "GEMINI CALIBRATION 2: EXPLOITING THE 1M+ CONTEXT WINDOW",
          "Unlike smaller models that truncate or compress instructions, Gemini 1.5 Pro / 2.0 / 3.8 can effortlessly generate 4,000+ words of granular curriculum in a single shot. The 'UNCOMPRESSED PRODUCTION DEPTH' directive forces Gemini to provide complete click-by-click SOPs without hand-waving or placeholders.",
          COLOR_SECONDARY
        ),
        new Paragraph({ spacing: { before: 100, after: 100 } }),
        createCallout(
          "GEMINI CALIBRATION 3: FORCING STRUCTURED MARKDOWN TABLES",
          "Gemini has class-leading reasoning on structured tabular datasets. Requiring Markdown tables for decision thresholds (Thumbstop %, Hold %, Outbound CTR, CPA, ROAS vs MER) ensures students get clean, visual SOP matrices rather than walls of text.",
          COLOR_ACCENT
        ),

        new Paragraph({ spacing: { before: 200, after: 100 } }),
        createParagraph("Generated for: Meta Ads Launchpad Course | Format: Editable Word (.docx) — Gemini AI Edition", { italics: true, color: COLOR_MUTED }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/home/agent/workspace/Meta_Ads_Launchpad_Course_Prompts.docx", buffer);
  console.log("Document successfully updated at /home/agent/workspace/Meta_Ads_Launchpad_Course_Prompts.docx");
});
