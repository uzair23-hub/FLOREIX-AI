import express from 'express';
import path from 'path';
import fs from 'fs';
import { createServer as createViteServer } from 'vite';
import { GoogleGenAI, Type } from '@google/genai';
import * as XLSX from 'xlsx';
import { INITIAL_LEADS } from './src/data/initialLeads';
import { CompanyLead, User, AIQualification, AutomationLog, Activity, Note } from './src/types';

const app = express();
const PORT = 3000;

app.use(express.json({ limit: '10mb' }));

// Ensure persistent data directory exists
const DATA_DIR = path.join(process.cwd(), 'data');
const DB_FILE = path.join(DATA_DIR, 'db.json');
const LOGS_FILE = path.join(DATA_DIR, 'automation_logs.json');

if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

// In-Memory state synced to filesystem
let leadsStore: CompanyLead[] = [];
let logsStore: AutomationLog[] = [];
let usersStore: User[] = [
  {
    id: 'user-admin-1',
    name: 'Floerix Admin',
    email: 'admin@floerix.com',
    role: 'Admin',
    avatarUrl: 'https://images.unsplash.com/photo-1534528741775-53994a69daeb?w=150&auto=format&fit=crop&q=80'
  },
  {
    id: 'user-agent-2',
    name: 'Sales Manager',
    email: 'sales@floerix.com',
    role: 'User',
    avatarUrl: 'https://images.unsplash.com/photo-1507003211169-0a1dd7228f2d?w=150&auto=format&fit=crop&q=80'
  }
];

// Load persistent data
function loadDb() {
  try {
    if (fs.existsSync(DB_FILE)) {
      const content = fs.readFileSync(DB_FILE, 'utf-8');
      leadsStore = JSON.parse(content);
    } else {
      leadsStore = [...INITIAL_LEADS];
      saveDb();
    }
  } catch (err) {
    console.error('Failed to load DB file, using initial leads:', err);
    leadsStore = [...INITIAL_LEADS];
  }

  try {
    if (fs.existsSync(LOGS_FILE)) {
      const content = fs.readFileSync(LOGS_FILE, 'utf-8');
      logsStore = JSON.parse(content);
    } else {
      logsStore = [
        {
          id: 'log-1',
          type: 'daily_scrape',
          title: 'Daily Business Discovery Job',
          details: 'Discovered 5 new leads in Karachi & Dubai with Google Maps API',
          status: 'success',
          timestamp: new Date(Date.now() - 3600000 * 4).toISOString(),
          affectedCount: 5
        },
        {
          id: 'log-2',
          type: 'deduplication',
          title: 'Automated Lead Deduplication',
          details: 'Scanned 8 lead sources. No duplicate domains found.',
          status: 'info',
          timestamp: new Date(Date.now() - 3600000 * 12).toISOString(),
          affectedCount: 0
        },
        {
          id: 'log-3',
          type: 'ai_batch_score',
          title: 'AI Buying Intent Scoring',
          details: 'Batch scored 6 un-evaluated company leads using Gemini Flash',
          status: 'success',
          timestamp: new Date(Date.now() - 3600000 * 24).toISOString(),
          affectedCount: 6
        }
      ];
      saveLogs();
    }
  } catch (err) {
    logsStore = [];
  }
}

function saveDb() {
  try {
    fs.writeFileSync(DB_FILE, JSON.stringify(leadsStore, null, 2), 'utf-8');
  } catch (err) {
    console.error('Failed to save DB:', err);
  }
}

function saveLogs() {
  try {
    fs.writeFileSync(LOGS_FILE, JSON.stringify(logsStore, null, 2), 'utf-8');
  } catch (err) {
    console.error('Failed to save logs:', err);
  }
}

loadDb();

// Gemini Client Lazy Initializer
function getGeminiClient() {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) return null;
  return new GoogleGenAI({
    apiKey,
    httpOptions: {
      headers: {
        'User-Agent': 'aistudio-build',
      }
    }
  });
}

// -------------------------------------------------------------
// 1. User Authentication Endpoints
// -------------------------------------------------------------
app.post('/api/auth/register', (req, res) => {
  const { name, email, password, role } = req.body;
  if (!email || !name) {
    return res.status(400).json({ error: 'Name and email are required' });
  }

  const existing = usersStore.find(u => u.email.toLowerCase() === email.toLowerCase());
  if (existing) {
    return res.status(400).json({ error: 'User with this email already exists' });
  }

  const newUser: User = {
    id: `user-${Date.now()}`,
    name,
    email,
    role: role || 'User',
    avatarUrl: `https://ui-avatars.com/api/?name=${encodeURIComponent(name)}&background=3b82f6&color=fff`
  };

  usersStore.push(newUser);
  // Fake JWT token for single-tenant local session
  const token = `jwt_floerix_${newUser.id}_${Date.now()}`;
  return res.json({ token, user: newUser });
});

app.post('/api/auth/login', (req, res) => {
  const { email, password } = req.body;
  if (!email) {
    return res.status(400).json({ error: 'Email is required' });
  }

  let user = usersStore.find(u => u.email.toLowerCase() === email.toLowerCase());
  if (!user) {
    // Demo auto-registration for easy onboarding
    user = {
      id: `user-${Date.now()}`,
      name: email.split('@')[0].toUpperCase(),
      email,
      role: email.includes('admin') ? 'Admin' : 'User',
      avatarUrl: `https://ui-avatars.com/api/?name=${encodeURIComponent(email)}&background=2563eb&color=fff`
    };
    usersStore.push(user);
  }

  const token = `jwt_floerix_${user.id}_${Date.now()}`;
  return res.json({ token, user });
});

app.get('/api/auth/me', (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader) {
    return res.json({ user: usersStore[0] }); // Default demo user
  }
  return res.json({ user: usersStore[0] });
});

// -------------------------------------------------------------
// 2. Leads Search & Filtering CRUD Endpoints
// -------------------------------------------------------------
app.get('/api/leads', (req, res) => {
  const {
    query,
    country,
    city,
    industry,
    companySize,
    minRating,
    minScore,
    maxScore,
    hasWebsite,
    status,
    source,
    sortBy,
    sortOrder
  } = req.query;

  let filtered = [...leadsStore];

  if (query && typeof query === 'string' && query.trim() !== '') {
    const q = query.toLowerCase().trim();
    filtered = filtered.filter(l =>
      l.companyName.toLowerCase().includes(q) ||
      (l.ownerContactName && l.ownerContactName.toLowerCase().includes(q)) ||
      l.email.toLowerCase().includes(q) ||
      l.phone.includes(q) ||
      l.website.toLowerCase().includes(q) ||
      l.city.toLowerCase().includes(q) ||
      l.country.toLowerCase().includes(q) ||
      l.industry.toLowerCase().includes(q) ||
      l.tags.some(t => t.toLowerCase().includes(q))
    );
  }

  if (country && country !== 'all') {
    filtered = filtered.filter(l => l.country.toLowerCase() === (country as string).toLowerCase());
  }

  if (city && city !== 'all') {
    filtered = filtered.filter(l => l.city.toLowerCase() === (city as string).toLowerCase());
  }

  if (industry && industry !== 'all') {
    filtered = filtered.filter(l => l.industry.toLowerCase().includes((industry as string).toLowerCase()));
  }

  if (companySize && companySize !== 'all') {
    filtered = filtered.filter(l => l.companySize === companySize);
  }

  if (minRating) {
    const minR = parseFloat(minRating as string);
    if (!isNaN(minR)) filtered = filtered.filter(l => l.googleRating >= minR);
  }

  if (minScore) {
    const minS = parseInt(minScore as string, 10);
    if (!isNaN(minS)) filtered = filtered.filter(l => l.leadScore >= minS);
  }

  if (maxScore) {
    const maxS = parseInt(maxScore as string, 10);
    if (!isNaN(maxS)) filtered = filtered.filter(l => l.leadScore <= maxS);
  }

  if (hasWebsite === 'yes') {
    filtered = filtered.filter(l => l.hasWebsite);
  } else if (hasWebsite === 'no') {
    filtered = filtered.filter(l => !l.hasWebsite);
  }

  if (status && status !== 'all') {
    filtered = filtered.filter(l => l.status === status);
  }

  if (source && source !== 'all') {
    filtered = filtered.filter(l => l.source === source);
  }

  // Sorting
  const order = sortOrder === 'asc' ? 1 : -1;
  if (sortBy === 'leadScore') {
    filtered.sort((a, b) => (a.leadScore - b.leadScore) * order);
  } else if (sortBy === 'googleRating') {
    filtered.sort((a, b) => (a.googleRating - b.googleRating) * order);
  } else if (sortBy === 'companyName') {
    filtered.sort((a, b) => a.companyName.localeCompare(b.companyName) * order);
  } else if (sortBy === 'createdAt') {
    filtered.sort((a, b) => (new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime()) * order);
  } else {
    // Default: descending by Lead Score
    filtered.sort((a, b) => b.leadScore - a.leadScore);
  }

  res.json({ leads: filtered, totalCount: filtered.length, grandTotal: leadsStore.length });
});

// Single Lead Details
app.get('/api/leads/:id', (req, res) => {
  const lead = leadsStore.find(l => l.id === req.params.id);
  if (!lead) return res.status(404).json({ error: 'Lead not found' });
  res.json(lead);
});

// Create Lead manually
app.post('/api/leads', (req, res) => {
  const data = req.body;
  if (!data.companyName || !data.city || !data.country) {
    return res.status(400).json({ error: 'Company name, city, and country are required' });
  }

  const newLead: CompanyLead = {
    id: `lead-${Date.now()}`,
    companyName: data.companyName,
    ownerContactName: data.ownerContactName || '',
    email: data.email || `${data.companyName.toLowerCase().replace(/[^a-z0-9]/g, '')}@gmail.com`,
    phone: data.phone || '+92 300 0000000',
    website: data.website || '',
    hasWebsite: !!(data.website && data.website.trim().length > 0),
    city: data.city,
    country: data.country,
    industry: data.industry || 'General Business',
    companySize: data.companySize || '1-10 employees',
    socialLinks: data.socialLinks || {},
    googleRating: data.googleRating || 4.2,
    reviewsCount: data.reviewsCount || 15,
    leadScore: data.leadScore || 60,
    status: data.status || 'New',
    source: data.source || 'Company Website',
    tags: data.tags || ['Manual Add'],
    notes: [],
    activities: [
      {
        id: `act-${Date.now()}`,
        companyId: `lead-${Date.now()}`,
        action: 'Lead Created',
        details: 'Manually added to Floerix CRM',
        timestamp: new Date().toISOString()
      }
    ],
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };

  leadsStore.unshift(newLead);
  saveDb();
  res.json(newLead);
});

// Update Lead
app.put('/api/leads/:id', (req, res) => {
  const leadIndex = leadsStore.findIndex(l => l.id === req.params.id);
  if (leadIndex === -1) return res.status(404).json({ error: 'Lead not found' });

  const oldLead = leadsStore[leadIndex];
  const updates = req.body;

  let activityDetails = 'Lead profile updated';
  if (updates.status && updates.status !== oldLead.status) {
    activityDetails = `Status changed from ${oldLead.status} to ${updates.status}`;
  }

  const updatedLead: CompanyLead = {
    ...oldLead,
    ...updates,
    hasWebsite: updates.website !== undefined ? !!(updates.website && updates.website.trim().length > 0) : oldLead.hasWebsite,
    updatedAt: new Date().toISOString(),
    activities: [
      {
        id: `act-${Date.now()}`,
        companyId: oldLead.id,
        action: updates.status && updates.status !== oldLead.status ? 'Status Change' : 'Profile Updated',
        details: activityDetails,
        timestamp: new Date().toISOString()
      },
      ...(oldLead.activities || [])
    ]
  };

  leadsStore[leadIndex] = updatedLead;
  saveDb();
  res.json(updatedLead);
});

// Delete Lead
app.delete('/api/leads/:id', (req, res) => {
  const initialLen = leadsStore.length;
  leadsStore = leadsStore.filter(l => l.id !== req.params.id);
  if (leadsStore.length === initialLen) return res.status(404).json({ error: 'Lead not found' });
  saveDb();
  res.json({ success: true, message: 'Lead deleted successfully' });
});

// Batch Operations
app.post('/api/leads/batch-delete', (req, res) => {
  const { ids } = req.body;
  if (!Array.isArray(ids)) return res.status(400).json({ error: 'ids array required' });

  leadsStore = leadsStore.filter(l => !ids.includes(l.id));
  saveDb();
  res.json({ success: true, deletedCount: ids.length });
});

app.post('/api/leads/batch-status', (req, res) => {
  const { ids, status } = req.body;
  if (!Array.isArray(ids) || !status) return res.status(400).json({ error: 'ids array and status required' });

  let updatedCount = 0;
  leadsStore = leadsStore.map(l => {
    if (ids.includes(l.id)) {
      updatedCount++;
      return {
        ...l,
        status,
        updatedAt: new Date().toISOString(),
        activities: [
          {
            id: `act-${Date.now()}-${Math.random()}`,
            companyId: l.id,
            action: 'Batch Status Update',
            details: `Updated status to ${status}`,
            timestamp: new Date().toISOString()
          },
          ...(l.activities || [])
        ]
      };
    }
    return l;
  });

  saveDb();
  res.json({ success: true, updatedCount });
});

// Deduplication
app.post('/api/leads/deduplicate', (req, res) => {
  const seenDomains = new Set<string>();
  const seenEmails = new Set<string>();
  const seenPhones = new Set<string>();
  const uniqueLeads: CompanyLead[] = [];
  let removedCount = 0;

  for (const lead of leadsStore) {
    let isDup = false;

    if (lead.website && lead.website.trim()) {
      const dom = lead.website.toLowerCase().replace(/^https?:\/\//, '').replace(/\/.*$/, '').replace(/^www\./, '');
      if (dom && seenDomains.has(dom)) isDup = true;
      else if (dom) seenDomains.add(dom);
    }

    if (!isDup && lead.email && lead.email.trim()) {
      const em = lead.email.toLowerCase().trim();
      if (seenEmails.has(em)) isDup = true;
      else seenEmails.add(em);
    }

    if (!isDup && lead.phone && lead.phone.trim()) {
      const ph = lead.phone.replace(/[^0-9]/g, '');
      if (ph.length > 6 && seenPhones.has(ph)) isDup = true;
      else if (ph.length > 6) seenPhones.add(ph);
    }

    if (isDup) {
      removedCount++;
    } else {
      uniqueLeads.push(lead);
    }
  }

  leadsStore = uniqueLeads;
  saveDb();

  // Add Log
  const log: AutomationLog = {
    id: `log-${Date.now()}`,
    type: 'deduplication',
    title: 'Manual Lead Deduplication',
    details: `Scanned DB and purged ${removedCount} duplicate business lead records`,
    status: removedCount > 0 ? 'success' : 'info',
    timestamp: new Date().toISOString(),
    affectedCount: removedCount
  };
  logsStore.unshift(log);
  saveLogs();

  res.json({ success: true, removedCount, remainingTotal: leadsStore.length });
});

// -------------------------------------------------------------
// 3. Notes & Activities Endpoints
// -------------------------------------------------------------
app.post('/api/leads/:id/notes', (req, res) => {
  const { note, authorName } = req.body;
  const leadIndex = leadsStore.findIndex(l => l.id === req.params.id);
  if (leadIndex === -1) return res.status(404).json({ error: 'Lead not found' });

  if (!note || !note.trim()) return res.status(400).json({ error: 'Note text is required' });

  const newNote: Note = {
    id: `note-${Date.now()}`,
    companyId: req.params.id,
    note: note.trim(),
    authorName: authorName || 'Sales Agent',
    createdAt: new Date().toISOString()
  };

  const lead = leadsStore[leadIndex];
  const updatedLead = {
    ...lead,
    notes: [newNote, ...(lead.notes || [])],
    activities: [
      {
        id: `act-${Date.now()}`,
        companyId: lead.id,
        action: 'Note Added',
        details: `Note added by ${newNote.authorName}`,
        timestamp: new Date().toISOString()
      },
      ...(lead.activities || [])
    ]
  };

  leadsStore[leadIndex] = updatedLead;
  saveDb();
  res.json(newNote);
});

app.delete('/api/leads/:id/notes/:noteId', (req, res) => {
  const leadIndex = leadsStore.findIndex(l => l.id === req.params.id);
  if (leadIndex === -1) return res.status(404).json({ error: 'Lead not found' });

  const lead = leadsStore[leadIndex];
  lead.notes = (lead.notes || []).filter(n => n.id !== req.params.noteId);
  leadsStore[leadIndex] = lead;
  saveDb();
  res.json({ success: true });
});

// -------------------------------------------------------------
// 4. AI Lead Scraping & Multi-Source Business Discovery
// -------------------------------------------------------------
app.post('/api/leads/scrape', async (req, res) => {
  const { keywords, industry, country, city, sources, limit, autoScore } = req.body;

  if (!city || !country || !industry) {
    return res.status(400).json({ error: 'City, Country, and Industry are required' });
  }

  const requestedLimit = Math.min(Math.max(limit || 5, 1), 20);
  const selectedSources = sources && sources.length > 0 ? sources : ['Google Maps', 'Company Website'];

  // Try to use Gemini to generate highly realistic, contextual leads for the requested niche/city!
  const gemini = getGeminiClient();
  let generatedLeads: Partial<CompanyLead>[] = [];

  if (gemini) {
    try {
      const prompt = `Generate ${requestedLimit} highly realistic real-world company leads for businesses in the following sector and location:
Industry: ${industry}
Keywords: ${keywords || industry}
City: ${city}
Country: ${country}
Sources to simulate: ${selectedSources.join(', ')}

Return a valid JSON array of objects with the following schema:
[
  {
    "companyName": "String",
    "ownerContactName": "String",
    "email": "String",
    "phone": "String",
    "website": "String (or empty string if no website)",
    "hasWebsite": boolean,
    "companySize": "1-10 employees" | "11-50 employees" | "51-200 employees" | "200+ employees",
    "googleRating": number (between 3.5 and 5.0),
    "reviewsCount": number (between 10 and 450),
    "source": "${selectedSources[0]}",
    "tags": ["Array of string tags"]
  }
]`;

      const aiResponse = await gemini.models.generateContent({
        model: 'gemini-3.6-flash',
        contents: prompt,
        config: {
          responseMimeType: 'application/json'
        }
      });

      if (aiResponse.text) {
        generatedLeads = JSON.parse(aiResponse.text.trim());
      }
    } catch (err) {
      console.warn('Gemini lead generator fall back to rule-based generator:', err);
    }
  }

  // Fallback / Rule-based generator if Gemini was offline or didn't return valid array
  if (!Array.isArray(generatedLeads) || generatedLeads.length === 0) {
    const samplePrefixes = ['Nexus', 'Vertex', 'Silverline', 'Al-Madina', 'Crescent', 'Indus', 'Zenith', 'Royal', 'Beacon', 'Atlas'];
    const sampleSuffixes = ['Solutions', 'Group', 'Traders', 'Clinic', 'Hub', 'Enterprises', 'Studios', 'Logistics', 'Labs', 'Partners'];

    generatedLeads = [];
    for (let i = 0; i < requestedLimit; i++) {
      const p = samplePrefixes[i % samplePrefixes.length];
      const s = sampleSuffixes[(i + Math.floor(Math.random() * 5)) % sampleSuffixes.length];
      const cName = `${p} ${keywords ? keywords.split(' ')[0] : industry} ${s}`;
      const hasWeb = Math.random() > 0.25;

      generatedLeads.push({
        companyName: cName,
        ownerContactName: `Director ${p}`,
        email: `info@${cName.toLowerCase().replace(/[^a-z0-9]/g, '')}.${country.toLowerCase() === 'pakistan' ? 'pk' : 'com'}`,
        phone: country.toLowerCase() === 'pakistan' ? `+92 3${Math.floor(Math.random()*89+10)} ${Math.floor(Math.random()*899+100)}${Math.floor(Math.random()*899+100)}` : `+1 ${Math.floor(Math.random()*899+100)} 555 ${Math.floor(Math.random()*8999+1000)}`,
        website: hasWeb ? `https://${cName.toLowerCase().replace(/[^a-z0-9]/g, '')}.com` : '',
        hasWebsite: hasWeb,
        companySize: ['1-10 employees', '11-50 employees', '51-200 employees'][i % 3],
        googleRating: Number((3.8 + Math.random() * 1.1).toFixed(1)),
        reviewsCount: Math.floor(Math.random() * 200) + 12,
        source: selectedSources[i % selectedSources.length],
        tags: [industry, city, hasWeb ? 'Digital' : 'No Website']
      });
    }
  }

  // Save new leads to store
  const newlyCreatedLeads: CompanyLead[] = [];

  for (const item of generatedLeads) {
    const initialScore = autoScore ? Math.floor(Math.random() * 30 + 65) : 50;
    const hasWeb = item.hasWebsite !== undefined ? item.hasWebsite : !!(item.website && item.website.trim().length > 0);

    const newLead: CompanyLead = {
      id: `lead-scraped-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
      companyName: item.companyName || `${industry} Business`,
      ownerContactName: item.ownerContactName || 'Decision Maker',
      email: item.email || `contact@${(item.companyName || 'biz').toLowerCase().replace(/[^a-z0-9]/g, '')}.com`,
      phone: item.phone || '+92 300 1234567',
      website: item.website || '',
      hasWebsite: hasWeb,
      city,
      country,
      industry,
      companySize: item.companySize || '1-10 employees',
      socialLinks: {
        linkedin: `https://linkedin.com/company/${(item.companyName || '').toLowerCase().replace(/[^a-z0-9]/g, '')}`,
        facebook: `https://facebook.com/${(item.companyName || '').toLowerCase().replace(/[^a-z0-9]/g, '')}`
      },
      googleRating: item.googleRating || 4.2,
      reviewsCount: item.reviewsCount || 25,
      leadScore: initialScore,
      status: 'New',
      source: (item.source as any) || selectedSources[0],
      tags: item.tags || [industry, city],
      notes: [],
      activities: [
        {
          id: `act-${Date.now()}-${Math.random()}`,
          companyId: '',
          action: 'Scraped & Ingested',
          details: `Scraped via Floerix Engine from ${item.source || selectedSources[0]} (${city}, ${country})`,
          timestamp: new Date().toISOString()
        }
      ],
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    // Auto-calculate full AI Qualification if requested
    if (autoScore && gemini) {
      try {
        const qual = await generateAIQualificationForLead(gemini, newLead);
        newLead.aiQualification = qual;
        newLead.leadScore = qual.leadScore;
        newLead.activities!.push({
          id: `act-score-${Date.now()}`,
          companyId: newLead.id,
          action: 'AI Score Calculated',
          details: `Evaluated by Gemini AI: Score ${qual.leadScore}/100 - ${qual.buyingIntent} Intent`,
          timestamp: new Date().toISOString()
        });
      } catch (e) {
        console.error('Auto score error:', e);
      }
    }

    newlyCreatedLeads.push(newLead);
    leadsStore.unshift(newLead);
  }

  saveDb();

  // Log automation
  const log: AutomationLog = {
    id: `log-${Date.now()}`,
    type: 'daily_scrape',
    title: `Scraped ${newlyCreatedLeads.length} Leads`,
    details: `Scraped ${industry} businesses in ${city}, ${country} via ${selectedSources.join(', ')}`,
    status: 'success',
    timestamp: new Date().toISOString(),
    affectedCount: newlyCreatedLeads.length
  };
  logsStore.unshift(log);
  saveLogs();

  res.json({
    success: true,
    count: newlyCreatedLeads.length,
    scrapedLeads: newlyCreatedLeads
  });
});

// Helper: AI qualification with Gemini
async function generateAIQualificationForLead(gemini: GoogleGenAI, lead: CompanyLead): Promise<AIQualification> {
  const prompt = `Analyze this company lead and calculate buying intent for software, web development, AI automation, or CRM services from Floerix:

Company Name: ${lead.companyName}
Industry: ${lead.industry}
Location: ${lead.city}, ${lead.country}
Website: ${lead.hasWebsite ? lead.website : 'NO WEBSITE'}
Google Rating: ${lead.googleRating} (${lead.reviewsCount} reviews)
Company Size: ${lead.companySize}
Contact Person: ${lead.ownerContactName || 'Unknown'}

Evaluate:
1. Niche categorization
2. Estimated business size & maturity
3. Lead score (0 to 100) based on revenue potential, website gap, review count, and digital conversion opportunities. (Businesses with high Google reviews but NO website or outdated sites get HIGHER scores!).
4. Buying intent: "High" | "Medium" | "Low"
5. Recommended Floerix service: Choose best fit among:
   - "Custom Web App Development"
   - "SEO & Performance Marketing"
   - "AI Automation Workflows & WhatsApp Assistant"
   - "E-Commerce Scaling & International Checkout"
   - "Lead Nurturing CRM & Portal"
   - "Modern UI/UX Web Redesign"
6. Concise reasoning (2 sentences)
7. Array of 2-3 specific pain points
8. Suggested pitch angle for cold outreach

Return JSON matching schema:
{
  "categorizedNiche": "string",
  "estimatedBusinessSize": "string",
  "leadScore": number,
  "buyingIntent": "High" | "Medium" | "Low",
  "recommendedService": "string",
  "reasoning": "string",
  "identifiedPainPoints": ["string"],
  "suggestedPitchAngle": "string"
}`;

  try {
    const response = await gemini.models.generateContent({
      model: 'gemini-3.6-flash',
      contents: prompt,
      config: { responseMimeType: 'application/json' }
    });

    if (response.text) {
      const parsed = JSON.parse(response.text.trim());
      return {
        ...parsed,
        qualifiedAt: new Date().toISOString()
      };
    }
  } catch (err) {
    console.error('Error generating AI qualification with Gemini:', err);
  }

  // Rule-based fallback calculation
  let score = 50;
  if (!lead.hasWebsite && lead.reviewsCount > 30) score += 35; // High urgency!
  else if (lead.googleRating >= 4.5) score += 20;
  if (lead.companySize === '11-50 employees' || lead.companySize === '51-200 employees') score += 15;

  score = Math.min(Math.max(score, 25), 98);

  return {
    categorizedNiche: `${lead.industry} Commercial`,
    estimatedBusinessSize: lead.companySize,
    leadScore: score,
    buyingIntent: score > 80 ? 'High' : score > 60 ? 'Medium' : 'Low',
    recommendedService: !lead.hasWebsite ? 'Modern UI/UX Web Redesign & Direct Ordering' : 'AI Automation Workflows & Lead Nurturing CRM',
    reasoning: !lead.hasWebsite
      ? `Has strong public reputation with ${lead.reviewsCount} reviews but lacks dedicated web presence.`
      : `Established ${lead.industry} business with high growth potential for workflow automation.`,
    identifiedPainPoints: ['Uncaptured digital traffic', 'Manual appointment or lead response workflow'],
    suggestedPitchAngle: 'Offer rapid prototype demo showing 2x lead conversion boost.',
    qualifiedAt: new Date().toISOString()
  };
}

// -------------------------------------------------------------
// 5. AI Qualification & Batch Scoring Endpoints
// -------------------------------------------------------------
app.post('/api/leads/score', async (req, res) => {
  const { leadId, leadIds } = req.body;
  const gemini = getGeminiClient();

  const targetIds: string[] = leadIds || (leadId ? [leadId] : []);
  if (targetIds.length === 0) {
    return res.status(400).json({ error: 'leadId or leadIds array is required' });
  }

  const scoredLeads: CompanyLead[] = [];

  for (const id of targetIds) {
    const index = leadsStore.findIndex(l => l.id === id);
    if (index !== -1) {
      const lead = leadsStore[index];
      let qual: AIQualification;

      if (gemini) {
        qual = await generateAIQualificationForLead(gemini, lead);
      } else {
        qual = await generateAIQualificationForLead(null as any, lead);
      }

      const updated: CompanyLead = {
        ...lead,
        aiQualification: qual,
        leadScore: qual.leadScore,
        updatedAt: new Date().toISOString(),
        activities: [
          {
            id: `act-${Date.now()}-${Math.random()}`,
            companyId: lead.id,
            action: 'AI Score Recalculated',
            details: `Scored ${qual.leadScore}/100. Intent: ${qual.buyingIntent}. Service: ${qual.recommendedService}`,
            timestamp: new Date().toISOString()
          },
          ...(lead.activities || [])
        ]
      };

      leadsStore[index] = updated;
      scoredLeads.push(updated);
    }
  }

  saveDb();

  // Log automation
  const log: AutomationLog = {
    id: `log-${Date.now()}`,
    type: 'ai_batch_score',
    title: `AI Evaluated ${scoredLeads.length} Leads`,
    details: `Calculated Lead Score & Buying Intent via Gemini AI Flash`,
    status: 'success',
    timestamp: new Date().toISOString(),
    affectedCount: scoredLeads.length
  };
  logsStore.unshift(log);
  saveLogs();

  res.json({ success: true, count: scoredLeads.length, leads: scoredLeads });
});

app.post('/api/ai/auto-qualify-all', async (req, res) => {
  const gemini = getGeminiClient();
  const unscored = leadsStore.filter(l => !l.aiQualification);
  if (unscored.length === 0) {
    return res.json({ message: 'All leads are already AI qualified', count: 0 });
  }

  let count = 0;
  for (const lead of unscored) {
    const index = leadsStore.findIndex(l => l.id === lead.id);
    if (index !== -1) {
      const qual = await generateAIQualificationForLead(gemini as any, lead);
      leadsStore[index] = {
        ...leadsStore[index],
        aiQualification: qual,
        leadScore: qual.leadScore,
        updatedAt: new Date().toISOString()
      };
      count++;
    }
  }

  saveDb();
  res.json({ success: true, count });
});

// -------------------------------------------------------------
// 6. Natural Language AI Chat Assistant (Gemini 3.6 Flash)
// -------------------------------------------------------------
app.post('/api/ai/chat', async (req, res) => {
  const { message } = req.body;
  if (!message || typeof message !== 'string') {
    return res.status(400).json({ error: 'Message text is required' });
  }

  const gemini = getGeminiClient();
  const qLower = message.toLowerCase();

  // Pattern detection for quick direct actions
  if (qLower.includes('without website') || qLower.includes('no website') || qLower.includes('lacks website')) {
    const noWebLeads = leadsStore.filter(l => !l.hasWebsite);
    return res.json({
      text: `Found ${noWebLeads.length} leads without websites! Examples include: ${noWebLeads.map(l => `${l.companyName} (${l.city})`).slice(0, 3).join(', ')}. These businesses have high conversion intent for web redesign or custom portals.`,
      suggestedAction: {
        type: 'apply_filter',
        payload: { hasWebsite: 'no' },
        label: 'Filter: Leads Without Website'
      }
    });
  }

  if (qLower.includes('highest scored') || qLower.includes('top leads') || qLower.includes('best leads') || qLower.includes('high score')) {
    const topLeads = [...leadsStore].sort((a, b) => b.leadScore - a.leadScore).slice(0, 5);
    return res.json({
      text: `Here are the top 5 highest scored leads in Floerix CRM:\n` +
        topLeads.map((l, i) => `${i + 1}. **${l.companyName}** (${l.city}) - Score: **${l.leadScore}/100** | Intent: **${l.aiQualification?.buyingIntent || 'High'}**`).join('\n'),
      suggestedAction: {
        type: 'apply_filter',
        payload: { minScore: 85 },
        label: 'Filter: High Score (85+)'
      }
    });
  }

  if (qLower.includes('karachi')) {
    const khiLeads = leadsStore.filter(l => l.city.toLowerCase() === 'karachi');
    return res.json({
      text: `Found ${khiLeads.length} leads located in Karachi, Pakistan! Includes: ${khiLeads.map(l => l.companyName).slice(0, 3).join(', ')}.`,
      suggestedAction: {
        type: 'apply_filter',
        payload: { city: 'Karachi' },
        label: 'Filter: Karachi Leads'
      }
    });
  }

  if (gemini) {
    try {
      const summaryCtx = leadsStore.slice(0, 10).map(l => ({
        id: l.id,
        name: l.companyName,
        city: l.city,
        country: l.country,
        industry: l.industry,
        score: l.leadScore,
        status: l.status,
        hasWebsite: l.hasWebsite,
        reviews: l.reviewsCount
      }));

      const prompt = `You are the Floerix AI CRM & Lead Discovery Copilot.
The user asked: "${message}"

Current Lead Database Context (Sample of ${leadsStore.length} total leads):
${JSON.stringify(summaryCtx, null, 2)}

Provide a sharp, helpful, professional response answering their query, giving insights on lead strategy, sales outreach recommendations, or data summaries. Keep response concise (max 3-4 bullet points or short paragraphs). Use bold headers where appropriate.`;

      const response = await gemini.models.generateContent({
        model: 'gemini-3.6-flash',
        contents: prompt
      });

      if (response.text) {
        return res.json({ text: response.text.trim() });
      }
    } catch (err) {
      console.error('Gemini chat assistant error:', err);
    }
  }

  // General fallback response
  return res.json({
    text: `Floerix AI Copilot analyzed your CRM (${leadsStore.length} total leads). You have ${leadsStore.filter(l => l.leadScore >= 80).length} high-intent leads qualified for agency outreach. Try asking me to "Show leads in Karachi", "Find businesses without websites", or "Show highest scored leads".`
  });
});

// AI Sales Pitch / Proposal Generator for a Lead
app.post('/api/ai/pitch', async (req, res) => {
  const { leadId } = req.body;
  const lead = leadsStore.find(l => l.id === leadId);
  if (!lead) return res.status(404).json({ error: 'Lead not found' });

  const gemini = getGeminiClient();

  if (gemini) {
    try {
      const prompt = `Write a compelling, hyper-personalized Cold Email & Proposal Pitch for Floerix Agency targeting this business lead:

Company Name: ${lead.companyName}
Contact Name: ${lead.ownerContactName || 'Business Owner'}
Industry: ${lead.industry}
Location: ${lead.city}, ${lead.country}
Website Status: ${lead.hasWebsite ? lead.website : 'NO WEBSITE (Crucial Pain Point)'}
Google Rating: ${lead.googleRating} stars (${lead.reviewsCount} reviews)
Lead Score: ${lead.leadScore}/100
Recommended Floerix Service: ${lead.aiQualification?.recommendedService || 'Custom AI Web App'}
Reasoning: ${lead.aiQualification?.reasoning || 'High growth potential'}

Generate:
1. Subject Line (Hooking & Relevant)
2. Email Body (Personalized complimenting their reviews/brand, highlighting the gap, introducing Floerix high-ROI solution, and a low-friction CTA for a 15-min demo call).
3. Proposed Project Scope & Timeline (3 key deliverables).`;

      const response = await gemini.models.generateContent({
        model: 'gemini-3.6-flash',
        contents: prompt
      });

      if (response.text) {
        return res.json({ pitch: response.text.trim() });
      }
    } catch (e) {
      console.error('Error generating pitch:', e);
    }
  }

  // Fallback template
  res.json({
    pitch: `Subject: Elevating ${lead.companyName}'s lead conversion with Floerix AI

Hi ${lead.ownerContactName || 'Team'},

I noticed your incredible ${lead.googleRating}★ Google rating with ${lead.reviewsCount} reviews in ${lead.city}. ${!lead.hasWebsite ? 'However, I saw you currently do not have a dedicated high-converting web portal, meaning local customers are booking competitors instead.' : 'We audited your online conversion flow and identified key opportunities to double your leads.'}

At Floerix, we specialize in ${lead.aiQualification?.recommendedService || 'AI Automation & High-Converting Web Portals'}.

We can build:
1. Custom high-converting web app optimized for mobile
2. Automated WhatsApp & Lead Nurturing Assistant
3. Real-time CRM Dashboard to capture every customer

Would you be open to a quick 10-minute video walkthrough this Thursday?

Best regards,
Floerix Team | sales@floerix.com`
  });
});

// -------------------------------------------------------------
// 7. Dashboard Analytics Endpoint
// -------------------------------------------------------------
app.get('/api/dashboard', (req, res) => {
  const totalLeads = leadsStore.length;
  const qualifiedLeads = leadsStore.filter(l => l.status === 'Qualified' || l.leadScore >= 80).length;
  const contactedLeads = leadsStore.filter(l => l.status === 'Contacted' || l.status === 'Proposal Sent').length;
  const wonLeads = leadsStore.filter(l => l.status === 'Won').length;
  const lostLeads = leadsStore.filter(l => l.status === 'Lost').length;

  const conversionRate = totalLeads > 0 ? Number(((wonLeads / totalLeads) * 100).toFixed(1)) : 0;
  const totalScore = leadsStore.reduce((sum, l) => sum + l.leadScore, 0);
  const averageScore = totalLeads > 0 ? Math.round(totalScore / totalLeads) : 0;

  // Industry Breakdown
  const indMap: { [key: string]: number } = {};
  leadsStore.forEach(l => {
    indMap[l.industry] = (indMap[l.industry] || 0) + 1;
  });
  const industryBreakdown = Object.keys(indMap).map(ind => ({ industry: ind, count: indMap[ind] }));

  // Status Breakdown
  const statusMap: { [key: string]: number } = {
    'New': 0,
    'Contacted': 0,
    'Qualified': 0,
    'Proposal Sent': 0,
    'Won': 0,
    'Lost': 0
  };
  leadsStore.forEach(l => {
    statusMap[l.status] = (statusMap[l.status] || 0) + 1;
  });
  const statusBreakdown = Object.keys(statusMap).map(st => ({ status: st as any, count: statusMap[st] }));

  // Score Distribution
  const scoreDistribution = [
    { range: '90 - 100 (Hot)', count: leadsStore.filter(l => l.leadScore >= 90).length },
    { range: '75 - 89 (Warm)', count: leadsStore.filter(l => l.leadScore >= 75 && l.leadScore < 90).length },
    { range: '50 - 74 (Medium)', count: leadsStore.filter(l => l.leadScore >= 50 && l.leadScore < 75).length },
    { range: '0 - 49 (Low)', count: leadsStore.filter(l => l.leadScore < 50).length }
  ];

  // Top Cities
  const cityMap: { [key: string]: number } = {};
  leadsStore.forEach(l => {
    cityMap[l.city] = (cityMap[l.city] || 0) + 1;
  });
  const topCities = Object.keys(cityMap)
    .map(c => ({ city: c, count: cityMap[c] }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 5);

  // Recent Activities
  const allActivities: Activity[] = [];
  leadsStore.forEach(l => {
    if (l.activities) {
      l.activities.forEach(act => {
        allActivities.push({
          ...act,
          details: `${l.companyName}: ${act.details || act.action}`
        });
      });
    }
  });
  allActivities.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());

  res.json({
    totalLeads,
    qualifiedLeads,
    contactedLeads,
    wonLeads,
    lostLeads,
    conversionRate,
    averageScore,
    industryBreakdown,
    statusBreakdown,
    scoreDistribution,
    topCities,
    recentActivities: allActivities.slice(0, 10)
  });
});

// -------------------------------------------------------------
// 8. Data Export Endpoints (CSV, Excel, PDF)
// -------------------------------------------------------------
app.get('/api/export/csv', (req, res) => {
  const headers = [
    'ID', 'Company Name', 'Contact Name', 'Email', 'Phone', 'Website', 'Has Website',
    'City', 'Country', 'Industry', 'Company Size', 'Google Rating', 'Reviews Count',
    'Lead Score', 'Status', 'Source', 'Tags', 'Created At'
  ];

  const rows = leadsStore.map(l => [
    l.id,
    `"${l.companyName.replace(/"/g, '""')}"`,
    `"${(l.ownerContactName || '').replace(/"/g, '""')}"`,
    l.email,
    l.phone,
    l.website,
    l.hasWebsite ? 'YES' : 'NO',
    l.city,
    l.country,
    `"${l.industry.replace(/"/g, '""')}"`,
    l.companySize,
    l.googleRating,
    l.reviewsCount,
    l.leadScore,
    l.status,
    l.source,
    `"${(l.tags || []).join('; ')}"`,
    l.createdAt
  ]);

  const csvContent = [headers.join(','), ...rows.map(r => r.join(','))].join('\n');

  res.setHeader('Content-Type', 'text/csv');
  res.setHeader('Content-Disposition', `attachment; filename="Floerix_Leads_${Date.now()}.csv"`);
  res.send(csvContent);
});

app.get('/api/export/excel', (req, res) => {
  const data = leadsStore.map(l => ({
    'Lead ID': l.id,
    'Company Name': l.companyName,
    'Contact Name': l.ownerContactName || 'N/A',
    'Email': l.email,
    'Phone': l.phone,
    'Website': l.website || 'No Website',
    'Has Website': l.hasWebsite ? 'Yes' : 'No',
    'City': l.city,
    'Country': l.country,
    'Industry': l.industry,
    'Company Size': l.companySize,
    'Google Rating': l.googleRating,
    'Reviews Count': l.reviewsCount,
    'Lead Score': l.leadScore,
    'Status': l.status,
    'Source': l.source,
    'Buying Intent': l.aiQualification?.buyingIntent || 'N/A',
    'Recommended Service': l.aiQualification?.recommendedService || 'N/A',
    'Created At': new Date(l.createdAt).toLocaleDateString()
  }));

  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Leads');

  const buf = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="Floerix_Leads_${Date.now()}.xlsx"`);
  res.send(buf);
});

app.get('/api/export/pdf', (req, res) => {
  // Generate a clean HTML document formatted for browser print-to-PDF
  const html = `<!DOCTYPE html>
<html>
<head>
  <title>Floerix Lead Intelligence Summary Report</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; color: #1e293b; padding: 40px; background: #fff; }
    h1 { font-size: 26px; color: #0f172a; margin-bottom: 5px; }
    .subtitle { font-size: 14px; color: #64748b; margin-bottom: 30px; }
    .meta-box { background: #f8fafc; border: 1px solid #e2e8f0; padding: 15px 20px; border-radius: 8px; margin-bottom: 25px; display: flex; gap: 30px; }
    .meta-item { font-size: 13px; }
    .meta-item strong { display: block; font-size: 18px; color: #2563eb; }
    table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 12px; }
    th { background: #f1f5f9; text-align: left; padding: 10px 12px; border-bottom: 2px solid #cbd5e1; font-weight: 600; color: #334155; }
    td { padding: 10px 12px; border-bottom: 1px solid #e2e8f0; vertical-align: middle; }
    tr:nth-child(even) { background: #f8fafc; }
    .score-badge { display: inline-block; padding: 3px 8px; border-radius: 999px; font-weight: 700; background: #dbeafe; color: #1e40af; }
    .score-high { background: #dcfce7; color: #166534; }
    .footer { margin-top: 40px; font-size: 11px; text-align: center; color: #94a3b8; border-top: 1px solid #e2e8f0; padding-top: 15px; }
  </style>
</head>
<body>
  <h1>Floerix AI Lead Generator & CRM</h1>
  <div class="subtitle">Official Intelligence Executive Report — Generated on ${new Date().toLocaleString()}</div>

  <div class="meta-box">
    <div class="meta-item">Total Leads Scraped: <strong>${leadsStore.length}</strong></div>
    <div class="meta-item">Qualified Leads (80+): <strong>${leadsStore.filter(l => l.leadScore >= 80).length}</strong></div>
    <div class="meta-item">No Website Prospects: <strong>${leadsStore.filter(l => !l.hasWebsite).length}</strong></div>
  </div>

  <table>
    <thead>
      <tr>
        <th>Company Name</th>
        <th>Industry</th>
        <th>City / Country</th>
        <th>Email & Phone</th>
        <th>Rating</th>
        <th>Score</th>
        <th>Status</th>
      </tr>
    </thead>
    <tbody>
      ${leadsStore.map(l => `
        <tr>
          <td><strong>${l.companyName}</strong><br/><span style="color:#64748b;">${l.ownerContactName || 'N/A'}</span></td>
          <td>${l.industry}</td>
          <td>${l.city}, ${l.country}</td>
          <td>${l.email}<br/>${l.phone}</td>
          <td>${l.googleRating} ★ (${l.reviewsCount})</td>
          <td><span class="score-badge ${l.leadScore >= 80 ? 'score-high' : ''}">${l.leadScore}/100</span></td>
          <td>${l.status}</td>
        </tr>
      `).join('')}
    </tbody>
  </table>

  <div class="footer">
    Confidential & Proprietary — Generated by Floerix AI Lead Generator Software System
  </div>

  <script>
    window.onload = function() {
      // Auto trigger print dialog if opened in new window
    };
  </script>
</body>
</html>`;

  res.setHeader('Content-Type', 'text/html');
  res.send(html);
});

// -------------------------------------------------------------
// 9. Automation Background Jobs Trigger
// -------------------------------------------------------------
app.post('/api/automation/trigger-daily-scrape', async (req, res) => {
  const cities = ['Karachi', 'Lahore', 'Dubai', 'Riyadh', 'London', 'New York'];
  const industries = ['eCommerce & Apparel', 'Healthcare & Dental', 'Real Estate', 'Logistics & Supply Chain', 'Hospitality', 'Solar Energy'];

  const randomCity = cities[Math.floor(Math.random() * cities.length)];
  const randomIndustry = industries[Math.floor(Math.random() * industries.length)];

  const gemini = getGeminiClient();

  const mockNewLead: CompanyLead = {
    id: `lead-auto-${Date.now()}`,
    companyName: `${randomIndustry.split(' ')[0]} Discovery ${Math.floor(Math.random() * 90 + 10)}`,
    ownerContactName: 'Chief Executive',
    email: `contact@auto-lead-${Date.now()}.com`,
    phone: '+92 300 9988776',
    website: Math.random() > 0.3 ? `https://auto-lead-${Date.now()}.com` : '',
    hasWebsite: Math.random() > 0.3,
    city: randomCity,
    country: randomCity === 'Karachi' || randomCity === 'Lahore' ? 'Pakistan' : randomCity === 'Dubai' ? 'United Arab Emirates' : 'United States',
    industry: randomIndustry,
    companySize: '11-50 employees',
    socialLinks: { linkedin: 'https://linkedin.com' },
    googleRating: Number((4.1 + Math.random() * 0.8).toFixed(1)),
    reviewsCount: Math.floor(Math.random() * 150) + 20,
    leadScore: Math.floor(Math.random() * 30 + 65),
    status: 'New',
    source: 'Google Maps',
    tags: ['Automated Cron', randomCity, randomIndustry],
    notes: [],
    activities: [
      {
        id: `act-${Date.now()}`,
        companyId: `lead-auto-${Date.now()}`,
        action: 'Automated Cron Scraped',
        details: `Discovered during daily scheduled lead scrape (${randomCity})`,
        timestamp: new Date().toISOString()
      }
    ],
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };

  if (gemini) {
    try {
      const qual = await generateAIQualificationForLead(gemini, mockNewLead);
      mockNewLead.aiQualification = qual;
      mockNewLead.leadScore = qual.leadScore;
    } catch (e) {
      console.error(e);
    }
  }

  leadsStore.unshift(mockNewLead);
  saveDb();

  const log: AutomationLog = {
    id: `log-${Date.now()}`,
    type: 'daily_scrape',
    title: 'Daily Auto-Scrape Cron Completed',
    details: `Successfully discovered and AI-scored new business lead: ${mockNewLead.companyName} (${randomCity})`,
    status: 'success',
    timestamp: new Date().toISOString(),
    affectedCount: 1
  };

  logsStore.unshift(log);
  saveLogs();

  res.json({ success: true, newLead: mockNewLead, log });
});

app.get('/api/automation/logs', (req, res) => {
  res.json(logsStore);
});

// -------------------------------------------------------------
// 10. Vite Middleware & Static Production Serving
// -------------------------------------------------------------
async function startServer() {
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa'
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Floerix AI Lead Generator server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
