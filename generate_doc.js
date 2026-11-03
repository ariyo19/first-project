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
const COLOR_PRIMARY = "1E3A8A"; // Deep Blue
const COLOR_SECONDARY = "0D9488"; // Teal
const COLOR_DARK = "111827"; // Dark slate
const COLOR_MUTED = "4B5563"; // Muted gray
const COLOR_LIGHT_BG = "F3F4F6"; // Light gray bg
const COLOR_CODE_BG = "F8FAFC"; // Very light code bg
const COLOR_BORDER = "E5E7EB";

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

function createCallout(title, body) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      left: { style: BorderStyle.SINGLE, size: 24, color: COLOR_PRIMARY },
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
                    color: COLOR_PRIMARY,
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
      spacing: { before: 40, after: 40, line: 240 },
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
      left: { style: BorderStyle.SINGLE, size: 18, color: COLOR_SECONDARY },
      right: { style: BorderStyle.SINGLE, size: 6, color: "CBD5E1" },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            shading: { fill: COLOR_CODE_BG, type: ShadingType.CLEAR },
            margins: { top: 200, bottom: 200, left: 200, right: 200 },
            children: paragraphs,
          }),
        ],
      }),
    ],
  });
}

const masterPrompt = `[ROLE & CONTEXT]
Act as an elite Direct-Response Meta Ads Veteran with 20 years of hands-on media buying experience. You have personally managed over $50,000,000 in profitable paid social spend across e-commerce, info-products, high-ticket services, and SaaS. You survived every major algorithmic shift—from Right-Hand Rail ads and the 2016 Lookalike golden era to the iOS 14.5 attribution shock and today's AI-driven Andromeda/Advantage+ broad targeting reality.

You are engineering high-value, practitioner-grade curriculum modules for my signature course: "Meta Ads Launchpad".

[ANTI-AI STYLE DIRECTIVES - STRICTLY ENFORCED]
1. ZERO FLUFF & ZERO AI CLICHES: Never use generic buzzwords such as "in today's digital landscape", "game-changer", "unlock potential", "delve into", "skyrocket", "vital", or "leverage". If a sentence sounds like a generic agency blog post, delete it.
2. BRUTALLY SPECIFIC NUMBERS & THRESHOLDS: Always give concrete dollar amounts, metrics percentages, and statistical decision triggers (e.g., Thumbstop rate < 25%, Outbound CTR < 1.2%, Cost per Unique Add-to-Cart > 40% of target CPA, Kill if spend > 1.5x target CPA with 0 conversions).
3. IN-THE-TRENCHES JARGON: Speak like a media buyer in a Slack war-room. Use precise terminology: 3:2:2 Dynamic Creative Testing (DCT), Broad Targeting (No Interests), Conversions API (CAPI) Gateway, Event Quality Match Score, Thumbstop Rate vs Hold Rate, Outbound Click-Through Rate vs Link Click CTR, Blended MER (Marketing Efficiency Ratio) vs in-platform modeled ROAS, Vertical vs Horizontal scaling, Bid Caps/Cost Caps, and Attribution Lag.
4. ACTIONABLE SOPs & DIAGNOSTIC TREES: Structure content so a complete novice gets zero confusion, while a seasoned agency buyer would nod in respect. For every concept, provide the EXACT mechanical steps, the metric benchmark, and the failure-mode troubleshooting.

[COURSE TIMELINE SPECIFICATIONS]
Expand the following 4 milestones into comprehensive, practitioner-grade course modules:

---
MILESTONE 1: DAYS 1-3 | The Ironclad Foundation
"You will set up your Business Manager, configure your pixel, and establish your core targeting avatars. No more guesswork."
- Required Deep Dives:
  * Meta Business Portfolio architecture: Proper separation of personal profiles, 2FA security, Business Asset Groups, and payment fail-safes.
  * Bulletproof Tracking: Conversions API (CAPI) Gateway setup via Shopify/WordPress/Custom server, Aggregated Event Measurement, Event Match Quality (EMQ) targeting 8.0+/10, and first-party domain verification.
  * Avatar Angle Extraction: Moving beyond demographic fluff (e.g. "women 25-45") to Emotional Pain-State / Core Desires / Current Ineffective Solutions / Mechanism of Action. Provide 3 fill-in-the-blank Avatar Archetype templates.
  * Compliance Shield: The exact landing page, ad copy, and claim policies that trigger automated bans before you spend $1.

---
MILESTONE 2: WEEK 1 | The Direct-Response Creative Engine
"You launch your first Direct-Response creative tests. You will see initial data roll in and make your first few profitable sales."
- Required Deep Dives:
  * The 3:2:2 Dynamic Creative Testing (DCT) Architecture: 1 Ad Set, Dynamic Creative ON, 3 distinct visual hooks (video/static), 2 copy angles (short punchy vs long-form story), 2 headlines. Why DCT isolates variables without splitting budget.
  * Creative Anatomy of a Winner: Breakdown of the 3-Second Hook (Thumbstop), The Problem Pivot (seconds 3-10), The Mechanism Demonstration (seconds 10-30), and The Irresistible CTA with Offer Stacking.
  * Setting Daily Test Budgets: Exact formula to set testing budget based on Target Cost Per Acquisition (CPA) (e.g., 3x-5x target CPA per ad set over 72 hours).
  * Day 1-7 Reading the Diagnostics Dashboard: Custom Meta Ads column layout setup (Thumbstop %, Hold %, Outbound CTR, CPC Outbound, Cost per Content View, Cost per Add-to-Cart, Purchases, ROAS). The exact decision table for first profitable sales.

---
MILESTONE 3: WEEK 2 | The Pruning & Scaling Logic
"We kill the losers and scale the winners. You'll master retargeting logic, bringing back the people who clicked but didn't buy."
- Required Deep Dives:
  * The No-Emotion Kill Decision Tree: Exact rule-sets for killing ads at 0.5x, 1x, and 1.5x Target CPA. How to avoid premature pausing before statistical significance (impressions & conversion lag).
  * Modern Retargeting Truth (Post-iOS14): Why hyper-fragmented 30-day website visitor retargeting fails today. How to build:
    1) High-Intent MoFu/BoFu retargeting (7-day Add to Cart / Initiate Checkout exclusions),
    2) In-Platform Retargeting (Instagram/Facebook Engagers 90D, Video Viewers 50%+),
    3) Advantage+ Catalog Ads (DPA) dynamic retargeting.
  * Winner Graduation Protocol: Extracting the winning Post ID (Social Proof preservation) from the DCT and migrating it to a CBO (Campaign Budget Optimization) Scaling Campaign.

---
MILESTONE 4: MONTH 1 | The Scalable Performance Media Machine
"You achieve consistent, predictable daily sales. Your ROAS stabilizes, your unit economics are heavily in the green, and you finally have a scalable performance media machine."
- Required Deep Dives:
  * Unit Economics & The Profit Ledger: Blended MER (Total Revenue / Total Ad Spend) vs in-platform reported ROAS. Calculating Break-Even CPA, Target Contribution Margin, and Lifetime Value (LTV) buffer.
  * Dual Scaling Protocols:
    - Protocol A (Vertical Scaling): The 15%-20% budget increase rule every 48-72 hours without resetting learning phase.
    - Protocol B (Horizontal Scaling): Testing fresh creative angles, broad seed testing, cost cap/bid cap scaling campaigns for high-volume days.
  * Creative Fatigue Management System: Setting up a weekly creative pipeline (2-4 new concepts per week) so performance never drops off a cliff.
  * Long-Term Stability Checklist: Operating the account in 30 minutes a day without tweaking compulsively.

[OUTPUT FORMAT REQUIRED]
Format each milestone as an exhaustive, step-by-step master training guide with:
1. Executive Summary & Core Philosophy (1 paragraph, punchy, pragmatic)
2. Exact Step-by-Step Operating Procedures (Actionable, chronological)
3. Data Benchmarks & Hard Decision Matrix (Tables with exact metrics and actions)
4. War-Room Warning: The #1 Beginner Mistake that bleeds cash at this stage
5. Student Action Homework & Deliverable Checklist`;

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
              text: "Master Prompt Architecture & Implementation Guide (20-Year Media Buyer Edition)",
              italics: true,
              size: 24,
              color: COLOR_MUTED,
              font: "Calibri",
            }),
          ],
          spacing: { before: 0, after: 280 },
        }),

        createCallout(
          "DOCUMENT OVERVIEW & PURPOSE",
          "This document provides the high-octane Claude mega-prompt designed by a 20-year direct response media buyer. It forces Claude to strip away generic AI marketing fluff and produce institutional-grade, battle-tested curriculum modules for your 'Meta Ads Launchpad Course' across all 4 key transformation milestones."
        ),

        new Paragraph({ spacing: { before: 200, after: 100 } }),

        // SECTION 1
        createHeaderP("SECTION 1: THE MASTER CLAUDE MEGA-PROMPT", HeadingLevel.HEADING_1),
        createParagraph(
          "Copy and paste the entire prompt box below directly into Claude (works best with Claude 3.5 Sonnet / Claude 3 Opus). It includes strict anti-AI constraints, operational metric thresholds, and exhaustive module specifications.",
          { italics: true, color: COLOR_MUTED }
        ),
        new Paragraph({ spacing: { before: 80, after: 120 } }),

        createPromptBox(masterPrompt),

        new Paragraph({ spacing: { before: 300, after: 100 } }),

        // SECTION 2
        createHeaderP("SECTION 2: MODULAR DEEP-DIVE PROMPTS (PER TIMELINE STAGE)", HeadingLevel.HEADING_1),
        createParagraph(
          "If you prefer to generate your course lessons milestone by milestone for maximum depth, use the targeted prompts below.",
          { color: COLOR_MUTED }
        ),

        // Phase 1
        createHeaderP("Milestone 1: Days 1-3 | Business Manager, Pixel & Avatar Mapping", HeadingLevel.HEADING_2),
        createParagraph(
          "Input Milestone: Days 1-3: You will set up your Business Manager, configure your pixel, and establish your core targeting avatars. No more guesswork.",
          { bold: true }
        ),
        createPromptBox(`Act as a 20-year performance media buyer who has spent $50M+ on Meta ads. I need you to write the complete, step-by-step curriculum for Module 1 of my "Meta Ads Launchpad Course" covering: Days 1-3: Setup Business Manager, configure Pixel/CAPI, and establish Core Targeting Avatars.

Guidelines:
1. Anti-AI tone: Zero generic fluff. Use real practitioner language and specific technical steps.
2. Provide the exact Business Portfolio hierarchy (Owners, 2FA, Business Asset Groups, Payment redundancy).
3. Conversions API (CAPI) and Pixel setup: Explain Event Match Quality (EMQ score 8.0+), Server-side tracking via CAPI Gateway, and Aggregated Event Measurement.
4. Avatar & Angle Blueprint: Give 3 fill-in-the-blank direct response avatar sheets: (A) The Frustrated Searcher, (B) The Skeptical Switcher, (C) The Aspirant. Detail their trigger events, internal dialogue, and hook angles.
5. Meta Compliance & Ban Prevention: The exact policy landmines that get accounts shut down on Day 1 (before/after claims, personal attributes, ungrounded ROI promises) and how to pass automated AI reviews.
6. Provide a student action checklist with exact pass/fail criteria.`),

        new Paragraph({ spacing: { before: 200, after: 100 } }),

        // Phase 2
        createHeaderP("Milestone 2: Week 1 | Direct-Response Creative Testing & First Profitable Sales", HeadingLevel.HEADING_2),
        createParagraph(
          "Input Milestone: Week 1: You launch your first Direct-Response creative tests. You will see initial data roll in and make your first few profitable sales.",
          { bold: true }
        ),
        createPromptBox(`Act as a 20-year direct response media buyer. Write the exhaustive curriculum for Week 1 of my Meta Ads Launchpad Course: "Direct-Response Creative Testing & Landing Your First Profitable Sales".

Include:
1. The 3:2:2 Dynamic Creative Testing (DCT) blueprint: Why we run 3 creatives, 2 copy angles, 2 headlines under 1 dynamic ad set. Explain why manual isolated ad sets waste budget and how Meta's machine learning distributes spend across DCT variables.
2. Hook Mastery & The 3-Second Rule:
   - Thumbstop rate benchmark (>25%-30%): How to calculate it (3-Second Video Views / Impressions).
   - Hold rate benchmark (>20% of thumbstops reaching 15s).
   - 5 plug-and-play hook frameworks for video and static creatives.
3. Budget Math: The formula to calculate testing budget based on target CPA (e.g. 3x-5x target CPA over 72 hours).
4. Meta Ads Reporting Dashboard Configuration: The exact custom column layout (Thumbstop %, Hold %, Outbound CTR, CPC Outbound, Cost per Unique Add to Cart, Purchases, ROAS).
5. Interpreting Day 1 to Day 7 data: A decision table showing what to do if you get high clicks but no sales, low clicks, or immediate profitable sales.`),

        new Paragraph({ spacing: { before: 200, after: 100 } }),

        // Phase 3
        createHeaderP("Milestone 3: Week 2 | The Kill/Scale Decision Tree & Retargeting Architecture", HeadingLevel.HEADING_2),
        createParagraph(
          "Input Milestone: Week 2: We kill the losers and scale the winners. You'll master retargeting logic, bringing back the people who clicked but didn't buy.",
          { bold: true }
        ),
        createPromptBox(`Act as a veteran 8-figure media buyer. Write the detailed training module for Week 2 of my Meta Ads Launchpad Course: "Kill The Losers, Scale The Winners, and Modern Retargeting Logic".

Provide:
1. The No-Emotion Kill Decision Tree:
   - When to kill an ad at 0.5x target CPA (Hook/CTR failure).
   - When to kill at 1.0x target CPA (Engagement without add-to-cart).
   - When to kill at 1.5x-2.0x target CPA with zero conversions.
   - How to handle attribution lag so students don't kill winning ads prematurely.
2. The Winner Extraction Protocol:
   - How to pull the winning Post ID from a Dynamic Creative ad to preserve social proof (likes, comments, shares).
   - How to graduate that Post ID into a dedicated scaling CBO campaign.
3. The Post-iOS14 Retargeting Architecture:
   - Explain why traditional 30-day website pixel retargeting is dead or hyper-expensive.
   - The 3-Tier Modern Retargeting Funnel:
     Tier 1: 7-day high-intent window (Add to Cart / Initiate Checkout, excluding 30-day purchasers).
     Tier 2: In-platform ecosystem retargeting (IG/FB profile engagers 60D, video viewers 50%+).
     Tier 3: Advantage+ Catalog Dynamic Product Ads (DPA).
4. Concrete copy & offer angles for retargeting: Objections handling, risk reversal, testimonials, and time-sensitive incentive.`),

        new Paragraph({ spacing: { before: 200, after: 100 } }),

        // Phase 4
        createHeaderP("Milestone 4: Month 1 | The Scalable Performance Media Machine", HeadingLevel.HEADING_2),
        createParagraph(
          "Input Milestone: Month 1: You achieve consistent, predictable daily sales. Your ROAS stabilizes, your unit economics are heavily in the green, and you finally have a scalable performance media machine.",
          { bold: true }
        ),
        createPromptBox(`Act as an elite media buyer with 20 years experience. Write the flagship Month 1 module for my Meta Ads Launchpad Course: "The Scalable Performance Media Machine & Unit Economics Stabilization".

Detail:
1. Unit Economics & Cash-in-Bank Math:
   - Blended MER (Marketing Efficiency Ratio) vs In-Platform modeled ROAS.
   - The Contribution Margin formula: Revenue minus COGS, ad spend, merchant fees, shipping/delivery costs.
   - Establishing the True Break-Even CPA and Target Scaling CPA.
2. The Two Primary Scaling Protocols:
   - Vertical Scaling: The 15%-20% budget bump rule every 48-72 hours without throwing ad sets into learning reset.
   - Horizontal Scaling: Creative diversification (introducing new angles to prevent ad fatigue), broad interest expansion, and Cost Cap / Bid Cap safety-net campaigns.
3. Creative Fatigue Defense System:
   - Diagnosing creative fatigue (Frequency > 2.5, Outbound CTR dropping > 20%, CPA rising).
   - The Weekly Creative Pipeline: A simple system to produce 2-4 fresh hook variations weekly without burning out.
4. The 30-Minute Daily Account Management Routine:
   - What to check at 9:00 AM.
   - What to check at 6:00 PM.
   - The golden rule: Never touch an ad set during the first 6 hours of the day.`),

        new Paragraph({ spacing: { before: 300, after: 100 } }),

        // SECTION 3
        createHeaderP("SECTION 3: THE MEDIA BUYER'S RULEBOOK (ANTI-AI CALIBRATION)", HeadingLevel.HEADING_1),
        createParagraph(
          "Why standard prompts produce unusable content and how this prompt architecture guarantees elite quality:",
          { bold: true }
        ),
        createCallout(
          "RULE 1: METRICS OVER OPINIONS",
          "Generic AI says: 'Make sure your creative is engaging.'\nA 20-Year Media Buyer says: 'Your 3-second video hook must hit a 28%+ Thumbstop Rate. If it falls below 20%, your CPA will be 2.4x higher regardless of copy or targeting. Cut the first 1.5 seconds and zoom into the visual disruption.'"
        ),
        new Paragraph({ spacing: { before: 100, after: 100 } }),
        createCallout(
          "RULE 2: MODERN CONTEXT OVER 2017 TACTICS",
          "Generic AI says: 'Find 15 narrow interest groups and layer them.'\nA 20-Year Media Buyer says: 'Broad targeting with dynamic creatives wins because Meta's Andromeda retrieval algorithm treats the creative as the targeting. Layered interests cap scale and spike your CPMs.'"
        ),
        new Paragraph({ spacing: { before: 100, after: 100 } }),
        createCallout(
          "RULE 3: CASH-FLOW & UNIT ECONOMICS OVER ROAS VANITY",
          "Generic AI says: 'Aim for a 4.0 ROAS.'\nA 20-Year Media Buyer says: 'In-platform ROAS is modeled and delayed by 24-72 hours. Run on blended MER. If your gross margin is 65% and your target net profit is 20%, your max allowable acquisition cost is 45% of AOV. Track cash in bank every Monday morning.'"
        ),

        new Paragraph({ spacing: { before: 200, after: 100 } }),
        createParagraph("Generated for: Meta Ads Launchpad Course | Format: Editable Word (.docx)", { italics: true, color: COLOR_MUTED }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/home/agent/workspace/Meta_Ads_Launchpad_Course_Prompts.docx", buffer);
  console.log("Document successfully created at /home/agent/workspace/Meta_Ads_Launchpad_Course_Prompts.docx");
});
