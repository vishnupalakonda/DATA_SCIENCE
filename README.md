<!--
╔══════════════════════════════════════════════════════════════════╗
║       DATA INTELLECT // PREMIUM GITHUB README                   ║
║       Designed for: vishnupalakonda                             ║
║       Replace all [BRACKETED] placeholders before publishing    ║
╚══════════════════════════════════════════════════════════════════╝
-->

<!-- ═══════════════════════════════════════════════════════════════
     SECTION 1 — HERO BANNER
     Replace the SVG src with your own Canva/Figma-exported banner,
     or use this inline SVG directly — it renders on GitHub.
════════════════════════════════════════════════════════════════════ -->

<div align="center">

<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 900 220" width="100%">
  <defs>
    <!-- Dark slate background gradient -->
    <linearGradient id="bg" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%"   stop-color="#0d1117"/>
      <stop offset="50%"  stop-color="#161b22"/>
      <stop offset="100%" stop-color="#0d1117"/>
    </linearGradient>
    <!-- Cyan glow for nodes -->
    <radialGradient id="nodeGlow" cx="50%" cy="50%" r="50%">
      <stop offset="0%"   stop-color="#00ffe5" stop-opacity="0.9"/>
      <stop offset="100%" stop-color="#00ffe5" stop-opacity="0"/>
    </radialGradient>
    <!-- Edge glow filter -->
    <filter id="glow" x="-40%" y="-40%" width="180%" height="180%">
      <feGaussianBlur stdDeviation="3" result="coloredBlur"/>
      <feMerge><feMergeNode in="coloredBlur"/><feMergeNode in="SourceGraphic"/></feMerge>
    </filter>
    <filter id="softGlow" x="-20%" y="-20%" width="140%" height="140%">
      <feGaussianBlur stdDeviation="6" result="coloredBlur"/>
      <feMerge><feMergeNode in="coloredBlur"/><feMergeNode in="SourceGraphic"/></feMerge>
    </filter>
  </defs>

  <!-- Background -->
  <rect width="900" height="220" fill="url(#bg)" rx="10"/>

  <!-- Subtle grid lines -->
  <g stroke="#1e2a38" stroke-width="0.5" opacity="0.6">
    <line x1="0" y1="44" x2="900" y2="44"/>
    <line x1="0" y1="88" x2="900" y2="88"/>
    <line x1="0" y1="132" x2="900" y2="132"/>
    <line x1="0" y1="176" x2="900" y2="176"/>
    <line x1="90" y1="0" x2="90" y2="220"/>
    <line x1="180" y1="0" x2="180" y2="220"/>
    <line x1="270" y1="0" x2="270" y2="220"/>
    <line x1="360" y1="0" x2="360" y2="220"/>
    <line x1="450" y1="0" x2="450" y2="220"/>
  </g>

  <!-- Neural network edges (left side) -->
  <g stroke="#00ffe5" stroke-width="0.8" opacity="0.3" filter="url(#glow)">
    <!-- Layer 0 → Layer 1 -->
    <line x1="60"  y1="60"  x2="160" y2="50"/>
    <line x1="60"  y1="60"  x2="160" y2="110"/>
    <line x1="60"  y1="110" x2="160" y2="50"/>
    <line x1="60"  y1="110" x2="160" y2="110"/>
    <line x1="60"  y1="110" x2="160" y2="170"/>
    <line x1="60"  y1="160" x2="160" y2="110"/>
    <line x1="60"  y1="160" x2="160" y2="170"/>
    <!-- Layer 1 → Layer 2 -->
    <line x1="160" y1="50"  x2="260" y2="75"/>
    <line x1="160" y1="50"  x2="260" y2="145"/>
    <line x1="160" y1="110" x2="260" y2="75"/>
    <line x1="160" y1="110" x2="260" y2="145"/>
    <line x1="160" y1="170" x2="260" y2="75"/>
    <line x1="160" y1="170" x2="260" y2="145"/>
    <!-- Layer 2 → Layer 3 -->
    <line x1="260" y1="75"  x2="350" y2="110"/>
    <line x1="260" y1="145" x2="350" y2="110"/>
  </g>

  <!-- Bright animated-look edges -->
  <g stroke="#00ffe5" stroke-width="1.5" opacity="0.7" filter="url(#glow)">
    <line x1="60"  y1="110" x2="160" y2="110"/>
    <line x1="160" y1="110" x2="260" y2="75"/>
    <line x1="260" y1="75"  x2="350" y2="110"/>
  </g>

  <!-- Neural nodes — Layer 0 -->
  <g filter="url(#glow)">
    <circle cx="60" cy="60"  r="6" fill="#0d1117" stroke="#00ffe5" stroke-width="1.5"/>
    <circle cx="60" cy="110" r="7" fill="#00ffe5" opacity="0.9"/>
    <circle cx="60" cy="160" r="6" fill="#0d1117" stroke="#00ffe5" stroke-width="1.5"/>
  </g>
  <!-- Layer 1 -->
  <g filter="url(#glow)">
    <circle cx="160" cy="50"  r="5" fill="#0d1117" stroke="#00ffe5" stroke-width="1.2"/>
    <circle cx="160" cy="110" r="7" fill="#00ffe5" opacity="0.8"/>
    <circle cx="160" cy="170" r="5" fill="#0d1117" stroke="#00ffe5" stroke-width="1.2"/>
  </g>
  <!-- Layer 2 -->
  <g filter="url(#glow)">
    <circle cx="260" cy="75"  r="7" fill="#00ffe5" opacity="0.9"/>
    <circle cx="260" cy="145" r="5" fill="#0d1117" stroke="#00ffe5" stroke-width="1.2"/>
  </g>
  <!-- Layer 3 (output) -->
  <g filter="url(#softGlow)">
    <circle cx="350" cy="110" r="10" fill="#00ffe5" opacity="0.95"/>
    <circle cx="350" cy="110" r="18" fill="#00ffe5" opacity="0.12"/>
    <circle cx="350" cy="110" r="28" fill="#00ffe5" opacity="0.06"/>
  </g>

  <!-- Data pipeline dots flowing right -->
  <g fill="#00ffe5" opacity="0.5">
    <circle cx="390" cy="110" r="3"/>
    <circle cx="410" cy="110" r="2"/>
    <circle cx="425" cy="110" r="1.5"/>
  </g>

  <!-- Divider line -->
  <line x1="440" y1="20" x2="440" y2="200" stroke="#1e3a4a" stroke-width="1" opacity="0.8"/>

  <!-- RIGHT SIDE — Identity block -->
  <!-- Top label -->
  <text x="470" y="65" font-family="'Courier New', monospace" font-size="11"
        fill="#00ffe5" opacity="0.7" letter-spacing="4">DATA INTELLECT</text>

  <!-- Divider slash -->
  <text x="470" y="92" font-family="'Courier New', monospace" font-size="11"
        fill="#3a4a5a" letter-spacing="3">// ──────────────────</text>

  <!-- Main name -->
  <text x="468" y="135" font-family="Georgia, serif" font-size="34"
        fill="#e8f4f8" font-weight="700" letter-spacing="1">[YOUR NAME]</text>

  <!-- Role tagline -->
  <text x="470" y="162" font-family="'Courier New', monospace" font-size="12"
        fill="#4a7a8a" letter-spacing="2">DATA SCIENTIST  ·  ML ENGINEER</text>

  <!-- Status pill -->
  <rect x="470" y="178" width="130" height="22" rx="11" fill="#00ffe510" stroke="#00ffe5" stroke-width="0.8"/>
  <circle cx="485" cy="189" r="4" fill="#00ffe5" opacity="0.9"/>
  <text x="494" y="193" font-family="'Courier New', monospace" font-size="10"
        fill="#00ffe5" letter-spacing="1">OPEN TO WORK</text>

  <!-- Top-right corner tag -->
  <text x="840" y="28" font-family="'Courier New', monospace" font-size="9"
        fill="#1e3a4a" letter-spacing="2" text-anchor="middle">v2.0.26</text>
</svg>

</div>

<br/>

<!-- ═══════════════════════════════════════════════════════════════
     SECTION 2 — PERSONA SPLIT  (Bio + Key Metrics)
════════════════════════════════════════════════════════════════════ -->

<table width="100%" cellspacing="0" cellpadding="0" border="0">
<tr>

<!-- ── LEFT COL: Avatar + Bio + Badges ── -->
<td width="38%" valign="top" align="center">

<br/>

<!--
  Replace the src URL below with your actual GitHub avatar:
  https://avatars.githubusercontent.com/u/231016263?v=4
-->
<img
  src="https://avatars.githubusercontent.com/u/231016263?v=4"
  width="130"
  style="border-radius:50%; border: 2.5px solid #00ffe5;"
  alt="[YOUR NAME]"
/>

<br/><br/>

<samp>

**[YOUR NAME]**
<br/>
`Data Scientist · ML Engineer · Andhra Pradesh, India`

</samp>

<br/>

```
Predictive Analytics.
Deep Learning.
Data Architecture.
```

<br/>

<!-- Social badges — replace href links with your actual profiles -->
<a href="https://linkedin.com/in/[YOUR-LINKEDIN]">
  <img src="https://img.shields.io/badge/LinkedIn-0A66C2?style=flat-square&logo=linkedin&logoColor=white" alt="LinkedIn"/>
</a>
&nbsp;
<a href="https://kaggle.com/[YOUR-KAGGLE]">
  <img src="https://img.shields.io/badge/Kaggle-20BEFF?style=flat-square&logo=kaggle&logoColor=white" alt="Kaggle"/>
</a>
&nbsp;
<a href="https://[YOUR-PORTFOLIO-URL]">
  <img src="https://img.shields.io/badge/Portfolio-00FFE5?style=flat-square&logo=vercel&logoColor=black" alt="Portfolio"/>
</a>

<br/><br/>

</td>

<!-- ── SPACER ── -->
<td width="4%"></td>

<!-- ── RIGHT COL: Key Metrics Dashboard ── -->
<td width="58%" valign="top">

<br/>

<!--
  KEY METRICS CARD
  Built as an inline SVG so it renders on GitHub with no image hosting needed.
  Update the numbers to match your actual stats.
-->

<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 480 210" width="100%">
  <defs>
    <linearGradient id="cardBg" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%"  stop-color="#0d1117"/>
      <stop offset="100%" stop-color="#161b22"/>
    </linearGradient>
    <filter id="cardGlow">
      <feGaussianBlur stdDeviation="2" result="blur"/>
      <feMerge><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge>
    </filter>
  </defs>

  <!-- Card shell -->
  <rect width="480" height="210" rx="12" fill="url(#cardBg)"/>
  <rect width="480" height="210" rx="12" fill="none" stroke="#1e3a4a" stroke-width="1"/>
  <!-- Top accent bar -->
  <rect x="0" y="0" width="480" height="3" rx="2" fill="#00ffe5" opacity="0.9"/>

  <!-- Card header -->
  <text x="22" y="30" font-family="'Courier New', monospace" font-size="10"
        fill="#4a7a8a" letter-spacing="3">KEY METRICS  //  DASHBOARD</text>
  <line x1="22" y1="38" x2="458" y2="38" stroke="#1e3a4a" stroke-width="0.8"/>

  <!-- ── METRIC 1: Models Deployed ── -->
  <rect x="22" y="50" width="132" height="130" rx="8" fill="#0a1520" stroke="#1e3a4a" stroke-width="0.8"/>
  <rect x="22" y="50" width="132" height="3" rx="2" fill="#00ffe5" opacity="0.7"/>
  <text x="88" y="100" font-family="Georgia, serif" font-size="36"
        fill="#00ffe5" font-weight="700" text-anchor="middle" filter="url(#cardGlow)">14+</text>
  <text x="88" y="118" font-family="'Courier New', monospace" font-size="8.5"
        fill="#4a7a8a" letter-spacing="1" text-anchor="middle">MODELS</text>
  <text x="88" y="132" font-family="'Courier New', monospace" font-size="8.5"
        fill="#4a7a8a" letter-spacing="1" text-anchor="middle">DEPLOYED</text>
  <!-- Mini sparkline -->
  <polyline points="38,168 55,162 72,155 88,150 105,145 122,140 140,136"
            fill="none" stroke="#00ffe5" stroke-width="1.2" opacity="0.5"/>
  <circle cx="140" cy="136" r="2.5" fill="#00ffe5" opacity="0.8"/>

  <!-- ── METRIC 2: Pipelines Built ── -->
  <rect x="174" y="50" width="132" height="130" rx="8" fill="#0a1520" stroke="#1e3a4a" stroke-width="0.8"/>
  <rect x="174" y="50" width="132" height="3" rx="2" fill="#00c8aa" opacity="0.7"/>
  <text x="240" y="100" font-family="Georgia, serif" font-size="36"
        fill="#00c8aa" font-weight="700" text-anchor="middle" filter="url(#cardGlow)">08</text>
  <text x="240" y="118" font-family="'Courier New', monospace" font-size="8.5"
        fill="#4a7a8a" letter-spacing="1" text-anchor="middle">DATA PIPELINES</text>
  <text x="240" y="132" font-family="'Courier New', monospace" font-size="8.5"
        fill="#4a7a8a" letter-spacing="1" text-anchor="middle">BUILT</text>
  <!-- Mini bar chart -->
  <g fill="#00c8aa" opacity="0.5">
    <rect x="190" y="162" width="10" height="14" rx="2"/>
    <rect x="206" y="155" width="10" height="21" rx="2"/>
    <rect x="222" y="148" width="10" height="28" rx="2"/>
    <rect x="238" y="158" width="10" height="18" rx="2"/>
    <rect x="254" y="142" width="10" height="34" rx="2"/>
    <rect x="270" y="150" width="10" height="26" rx="2"/>
    <rect x="286" y="138" width="10" height="38" rx="2"/>
  </g>

  <!-- ── METRIC 3: Max Accuracy ── -->
  <rect x="326" y="50" width="132" height="130" rx="8" fill="#0a1520" stroke="#1e3a4a" stroke-width="0.8"/>
  <rect x="326" y="50" width="132" height="3" rx="2" fill="#7ee8d6" opacity="0.7"/>
  <text x="392" y="100" font-family="Georgia, serif" font-size="30"
        fill="#7ee8d6" font-weight="700" text-anchor="middle" filter="url(#cardGlow)">98.4%</text>
  <text x="392" y="118" font-family="'Courier New', monospace" font-size="8.5"
        fill="#4a7a8a" letter-spacing="1" text-anchor="middle">MAX MODEL</text>
  <text x="392" y="132" font-family="'Courier New', monospace" font-size="8.5"
        fill="#4a7a8a" letter-spacing="1" text-anchor="middle">ACCURACY</text>
  <!-- Radial arc progress -->
  <circle cx="392" cy="155" r="20" fill="none" stroke="#1e3a4a" stroke-width="3"/>
  <path d="M 392 135 A 20 20 0 1 1 375 171"
        fill="none" stroke="#7ee8d6" stroke-width="3" stroke-linecap="round" opacity="0.8"/>
  <circle cx="375" cy="171" r="3" fill="#7ee8d6" opacity="0.9"/>
</svg>

<br/><br/>

</td>
</tr>
</table>

<br/>

---

<!-- ═══════════════════════════════════════════════════════════════
     SECTION 3 — THE ARCHITECTURE (Data Science Workflow Matrix)
════════════════════════════════════════════════════════════════════ -->

<h3>
  <img src="https://img.shields.io/badge/─────────────────────────────────────────────────-0d1117?style=flat-square" alt=""/>
  &nbsp;⬡&nbsp; THE ARCHITECTURE &nbsp;·&nbsp; Data Science Workflow
</h3>

<table width="100%" cellspacing="0" cellpadding="0" border="0">
<tr>

<td width="25%" valign="top">

<!-- STAGE 1 -->
<table width="100%" cellspacing="0" cellpadding="8" style="border: 1px solid #1e3a4a; border-radius: 8px;">
<tr>
  <td align="center" bgcolor="#0d1117">
    <br/>
    <code style="color:#00ffe5; font-size:10px; letter-spacing:2px;">01 · INGESTION</code><br/>
    <sub style="color:#4a7a8a; letter-spacing:1px;">& PIPELINES</sub>
    <br/><br/>
    <img src="https://img.shields.io/badge/Apache Kafka-231F20?style=flat-square&logo=apachekafka&logoColor=white"/>
    <br/><br/>
    <img src="https://img.shields.io/badge/Apache Airflow-017CEE?style=flat-square&logo=apacheairflow&logoColor=white"/>
    <br/><br/>
    <img src="https://img.shields.io/badge/Apache Spark-E25A1C?style=flat-square&logo=apachespark&logoColor=white"/>
    <br/><br/>
  </td>
</tr>
</table>

</td>

<td width="1%" align="center" valign="middle">
  <br/><br/>
  <sub>→</sub>
</td>

<td width="25%" valign="top">

<!-- STAGE 2 -->
<table width="100%" cellspacing="0" cellpadding="8" style="border: 1px solid #1e3a4a; border-radius: 8px;">
<tr>
  <td align="center" bgcolor="#0d1117">
    <br/>
    <code style="color:#00c8aa; font-size:10px; letter-spacing:2px;">02 · STORAGE</code><br/>
    <sub style="color:#4a7a8a; letter-spacing:1px;">& ARCHITECTURE</sub>
    <br/><br/>
    <img src="https://img.shields.io/badge/PostgreSQL-4169E1?style=flat-square&logo=postgresql&logoColor=white"/>
    <br/><br/>
    <img src="https://img.shields.io/badge/AWS S3-FF9900?style=flat-square&logo=amazons3&logoColor=white"/>
    <br/><br/>
    <img src="https://img.shields.io/badge/Snowflake-29B5E8?style=flat-square&logo=snowflake&logoColor=white"/>
    <br/><br/>
  </td>
</tr>
</table>

</td>

<td width="1%" align="center" valign="middle">
  <br/><br/>
  <sub>→</sub>
</td>

<td width="25%" valign="top">

<!-- STAGE 3 -->
<table width="100%" cellspacing="0" cellpadding="8" style="border: 1px solid #1e3a4a; border-radius: 8px;">
<tr>
  <td align="center" bgcolor="#0d1117">
    <br/>
    <code style="color:#7ee8d6; font-size:10px; letter-spacing:2px;">03 · MODELING</code><br/>
    <sub style="color:#4a7a8a; letter-spacing:1px;">& COMPUTE</sub>
    <br/><br/>
    <img src="https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white"/>
    <br/><br/>
    <img src="https://img.shields.io/badge/PyTorch-EE4C2C?style=flat-square&logo=pytorch&logoColor=white"/>
    <br/><br/>
    <img src="https://img.shields.io/badge/Scikit--Learn-F7931E?style=flat-square&logo=scikit-learn&logoColor=white"/>
    <br/><br/>
  </td>
</tr>
</table>

</td>

<td width="1%" align="center" valign="middle">
  <br/><br/>
  <sub>→</sub>
</td>

<td width="23%" valign="top">

<!-- STAGE 4 -->
<table width="100%" cellspacing="0" cellpadding="8" style="border: 1px solid #1e3a4a; border-radius: 8px;">
<tr>
  <td align="center" bgcolor="#0d1117">
    <br/>
    <code style="color:#a8f0e4; font-size:10px; letter-spacing:2px;">04 · PRODUCTION</code><br/>
    <sub style="color:#4a7a8a; letter-spacing:1px;">& DEPLOYMENT</sub>
    <br/><br/>
    <img src="https://img.shields.io/badge/Docker-2496ED?style=flat-square&logo=docker&logoColor=white"/>
    <br/><br/>
    <img src="https://img.shields.io/badge/FastAPI-009688?style=flat-square&logo=fastapi&logoColor=white"/>
    <br/><br/>
    <img src="https://img.shields.io/badge/AWS EC2-FF9900?style=flat-square&logo=amazonec2&logoColor=white"/>
    <br/><br/>
  </td>
</tr>
</table>

</td>

</tr>
</table>

<br/>

---

<!-- ═══════════════════════════════════════════════════════════════
     SECTION 4 — INTERACTIVE DATA PRODUCTS (Project Showcase)
     Each card is a horizontal 2-column block.
     Drop your screenshot URL into the <img src=""> tags.
════════════════════════════════════════════════════════════════════ -->

<h3>&nbsp;⬡&nbsp; DATA PRODUCTS &nbsp;·&nbsp; Selected Work</h3>

<!-- ─── PROJECT CARD 1 ─── -->
<table width="100%" cellspacing="0" cellpadding="16" style="border: 1px solid #1e3a4a; border-radius: 10px; background: #0d1117;">
<tr>

  <!-- Screenshot placeholder -->
  <td width="38%" valign="middle" align="center">
    <!--
      Replace the src below with your actual project screenshot URL.
      Recommended: Upload the image to your GitHub repo and reference it.
      Example: src="https://raw.githubusercontent.com/vishnupalakonda/REPO/main/assets/project1.png"
    -->
    <img
      src="https://placehold.co/320x180/0d1117/00ffe5?text=[ Drop+Screenshot+Here ]&font=courier"
      width="100%"
      style="border-radius: 6px; border: 1px solid #1e3a4a;"
      alt="Project 1 Screenshot"
    />
  </td>

  <!-- Project details -->
  <td width="62%" valign="top">
    <br/>
    <code style="color:#00ffe5; font-size:10px; letter-spacing:3px;">PROJECT · 01</code>
    <br/><br/>
    <strong>[PROJECT TITLE ONE]</strong>
    <br/><br/>
    <sub>Business Impact: [One sentence describing the measurable business outcome — e.g., "Reduced customer churn by 23% by deploying a real-time propensity model serving 2M+ users."]</sub>
    <br/><br/>
    <!-- Stack tags -->
    <img src="https://img.shields.io/badge/[Tool 1]-0d1117?style=flat-square&labelColor=1e3a4a&color=00ffe520" alt="tag"/>
    &nbsp;
    <img src="https://img.shields.io/badge/[Tool 2]-0d1117?style=flat-square&labelColor=1e3a4a&color=00ffe520" alt="tag"/>
    &nbsp;
    <img src="https://img.shields.io/badge/[Tool 3]-0d1117?style=flat-square&labelColor=1e3a4a&color=00ffe520" alt="tag"/>
    <br/><br/>
    <a href="https://github.com/vishnupalakonda/[REPO-NAME]">
      <img src="https://img.shields.io/badge/View Repository →-00ffe510?style=flat-square&logo=github&logoColor=00ffe5&labelColor=0d1117&color=1e3a4a"/>
    </a>
    <br/>
  </td>

</tr>
</table>

<br/>

<!-- ─── PROJECT CARD 2 ─── -->
<table width="100%" cellspacing="0" cellpadding="16" style="border: 1px solid #1e3a4a; border-radius: 10px; background: #0d1117;">
<tr>

  <!-- Project details (flipped layout) -->
  <td width="62%" valign="top">
    <br/>
    <code style="color:#00c8aa; font-size:10px; letter-spacing:3px;">PROJECT · 02</code>
    <br/><br/>
    <strong>[PROJECT TITLE TWO]</strong>
    <br/><br/>
    <sub>Business Impact: [One sentence describing the measurable business outcome — e.g., "Automated an 8-step ETL pipeline cutting reporting latency from 6 hours to under 4 minutes."]</sub>
    <br/><br/>
    <!-- Stack tags -->
    <img src="https://img.shields.io/badge/[Tool 1]-0d1117?style=flat-square&labelColor=1e3a4a&color=00c8aa20" alt="tag"/>
    &nbsp;
    <img src="https://img.shields.io/badge/[Tool 2]-0d1117?style=flat-square&labelColor=1e3a4a&color=00c8aa20" alt="tag"/>
    &nbsp;
    <img src="https://img.shields.io/badge/[Tool 3]-0d1117?style=flat-square&labelColor=1e3a4a&color=00c8aa20" alt="tag"/>
    <br/><br/>
    <a href="https://github.com/vishnupalakonda/[REPO-NAME]">
      <img src="https://img.shields.io/badge/View Repository →-00c8aa10?style=flat-square&logo=github&logoColor=00c8aa&labelColor=0d1117&color=1e3a4a"/>
    </a>
    <br/>
  </td>

  <!-- Screenshot placeholder -->
  <td width="38%" valign="middle" align="center">
    <!--
      Replace the src below with your actual project screenshot URL.
    -->
    <img
      src="https://placehold.co/320x180/0d1117/00c8aa?text=[ Drop+Screenshot+Here ]&font=courier"
      width="100%"
      style="border-radius: 6px; border: 1px solid #1e3a4a;"
      alt="Project 2 Screenshot"
    />
  </td>

</tr>
</table>

<br/>

---

<!-- ═══════════════════════════════════════════════════════════════
     SECTION 5 — ANALYTICS FOOTER
     GitHub Readme Stats + GitHub Streak — all dark-themed with
     Cyber Cyan accents to match the banner palette.

     Replace "vishnupalakonda" with your actual GitHub username
     in every URL below if it differs.
════════════════════════════════════════════════════════════════════ -->

<h3>&nbsp;⬡&nbsp; SYSTEM METRICS &nbsp;·&nbsp; GitHub Analytics</h3>

<!-- Stats row -->
<div align="center">

<img
  src="https://github-readme-stats.vercel.app/api?username=vishnupalakonda&show_icons=true&theme=transparent&hide_border=true&title_color=00ffe5&icon_color=00ffe5&text_color=7ee8d6&bg_color=0d1117&ring_color=00ffe5&count_private=true&include_all_commits=true"
  height="165"
  alt="GitHub Stats"
/>
&nbsp;&nbsp;
<img
  src="https://github-readme-stats.vercel.app/api/top-langs/?username=vishnupalakonda&layout=compact&theme=transparent&hide_border=true&title_color=00ffe5&text_color=7ee8d6&bg_color=0d1117&langs_count=6"
  height="165"
  alt="Top Languages"
/>

<br/><br/>

<!-- Streak stats — full width -->
<img
  src="https://github-readme-streak-stats.herokuapp.com/?user=vishnupalakonda&theme=transparent&hide_border=true&ring=00ffe5&fire=00c8aa&currStreakLabel=00ffe5&sideLabels=7ee8d6&dates=4a7a8a&background=0d1117&stroke=1e3a4a"
  width="70%"
  alt="GitHub Streak"
/>

<br/><br/>

<!-- Activity graph — full width -->
<img
  src="https://github-readme-activity-graph.vercel.app/graph?username=vishnupalakonda&bg_color=0d1117&color=00ffe5&line=00c8aa&point=7ee8d6&area=true&area_color=00ffe5&hide_border=true&custom_title=Contribution%20Architecture"
  width="100%"
  alt="Activity Graph"
/>

<br/><br/>

<!-- Profile views counter + visitor badge -->
<img src="https://komarev.com/ghpvc/?username=vishnupalakonda&style=flat-square&color=00ffe5&label=PROFILE+VIEWS&labelColor=0d1117" alt="Profile Views"/>
&nbsp;&nbsp;
<img src="https://img.shields.io/github/followers/vishnupalakonda?style=flat-square&color=00ffe5&labelColor=0d1117&label=FOLLOWERS&logo=github&logoColor=00ffe5" alt="Followers"/>

</div>

<br/>

<!-- ── FOOTER RULE ── -->

<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 900 40" width="100%">
  <defs>
    <linearGradient id="ruleGrad" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%"   stop-color="#0d1117"/>
      <stop offset="20%"  stop-color="#00ffe5" stop-opacity="0.4"/>
      <stop offset="50%"  stop-color="#00ffe5" stop-opacity="0.9"/>
      <stop offset="80%"  stop-color="#00ffe5" stop-opacity="0.4"/>
      <stop offset="100%" stop-color="#0d1117"/>
    </linearGradient>
  </defs>
  <rect y="18" width="900" height="1" fill="url(#ruleGrad)"/>
  <text x="450" y="35" font-family="'Courier New', monospace" font-size="9"
        fill="#1e3a4a" letter-spacing="4" text-anchor="middle">DATA INTELLECT  ·  [YOUR NAME]  ·  [CURRENT YEAR]</text>
</svg>
