import React, { useEffect, useState, useRef } from "react";
import { jsPDF } from "jspdf";
import "jspdf-autotable";
import { saveAs } from "file-saver";
import * as XLSX from "xlsx";

// --- ICONS ---
const PencilIcon = ({ className }) => <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}><path strokeLinecap="round" strokeLinejoin="round" d="m16.862 4.487 1.687-1.688a1.875 1.875 0 1 1 2.652 2.652L10.582 16.07a4.5 4.5 0 0 1-1.897 1.13L6 18l.8-2.685a4.5 4.5 0 0 1 1.13-1.897l8.932-8.931Zm0 0L19.5 7.125M18 14v4.75A2.25 2.25 0 0 1 15.75 21H5.25A2.25 2.25 0 0 1 3 18.75V8.25A2.25 2.25 0 0 1 5.25 6H10" /></svg>;
const DownloadIcon = ({ className }) => <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}><path strokeLinecap="round" strokeLinejoin="round" d="M3 16.5v2.25A2.25 2.25 0 0 0 5.25 21h13.5A2.25 2.25 0 0 0 21 18.75V16.5M16.5 12 12 16.5m0 0L7.5 12m4.5 4.5V3" /></svg>;
const FileDocumentIcon = ({ className }) => <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}><path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 0 0-3.375-3.375h-1.5A1.125 1.125 0 0 1 13.5 7.125v-1.5a3.375 3.375 0 0 0-3.375-3.375H8.25m2.25 0H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 0 0-9-9Z" /></svg>;
const TableIcon = ({ className }) => <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}><path strokeLinecap="round" strokeLinejoin="round" d="M3.375 19.5h17.25m-17.25 0a1.125 1.125 0 0 1-1.125-1.125M3.375 19.5h7.5c.621 0 1.125-.504 1.125-1.125m-9.75 0V5.625m0 12.75v-1.5c0-.621.504-1.125 1.125-1.125m18.375 2.625V5.625m0 12.75c0 .621-.504 1.125-1.125 1.125m1.125-1.125v-1.5c0-.621-.504-1.125-1.125-1.125m0 3.75h-7.5A1.125 1.125 0 0 1 12 18.75m9.75-12.75c0-.621-.504-1.125-1.125-1.125H3.375c-.621 0-1.125.504-1.125 1.125m19.5 0v1.5c0 .621-.504 1.125-1.125 1.125M2.25 5.625v1.5c0 .621.504 1.125 1.125 1.125m0 0h17.25m-17.25 0h7.5c.621 0 1.125.504 1.125 1.125M3.375 8.25c-.621 0-1.125.504-1.125 1.125v1.5c0 .621.504 1.125 1.125 1.125m17.25-3.75h-7.5c-.621 0-1.125.504-1.125 1.125m8.625-1.125c.621 0 1.125.504 1.125 1.125v1.5c0 .621-.504 1.125-1.125 1.125m-17.25 0h7.5m-7.5 0c-.621 0-1.125.504-1.125 1.125v1.5c0 .621.504 1.125 1.125 1.125M12 10.875v-1.5m0 1.5c0 .621-.504 1.125-1.125 1.125M12 10.875c0 .621.504 1.125 1.125 1.125m-2.25 0c.621 0 1.125.504 1.125 1.125M13.125 12h7.5m-7.5 0c-.621 0-1.125.504-1.125 1.125M20.625 12c.621 0 1.125.504 1.125 1.125v1.5c0 .621-.504 1.125-1.125 1.125m-17.25 0h7.5M12 14.625v-1.5m0 1.5c0 .621-.504 1.125-1.125 1.125M12 14.625c0 .621.504 1.125 1.125 1.125m-2.25 0c.621 0 1.125.504 1.125 1.125m0 1.5v-1.5m0 0c0-.621.504-1.125 1.125-1.125m0 0h7.5" /></svg>;
const CalendarIcon = ({ className }) => <span className={className}>📅</span>;
const ProjectListIcon = ({ className }) => (
  <span className={`${className} relative flex items-center justify-center`}>
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className="w-full h-full text-yellow-500"><path strokeLinecap="round" strokeLinejoin="round" d="M6.75 3v2.25M17.25 3v2.25M3 18.75V7.5a2.25 2.25 0 0 1 2.25-2.25h13.5A2.25 2.25 0 0 1 21 7.5v11.25m-18 0A2.25 2.25 0 0 0 5.25 21h13.5A2.25 2.25 0 0 0 21 18.75m-18 0h18" /></svg>
  </span>
);
const Wrench = ({ className }) => <span className={className}>🔧</span>;
const EllipsisVerticalIcon = ({ className }) => <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}><path strokeLinecap="round" strokeLinejoin="round" d="M12 6.75a.75.75 0 1 1 0-1.5.75.75 0 0 1 0 1.5ZM12 12.75a.75.75 0 1 1 0-1.5.75.75 0 0 1 0 1.5ZM12 18.75a.75.75 0 1 1 0-1.5.75.75 0 0 1 0 1.5Z" /></svg>;
const EyeIcon = ({ className }) => <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}><path strokeLinecap="round" strokeLinejoin="round" d="M2.036 12.322a1.012 1.012 0 0 1 0-.639l4.25-6.5a1.012 1.012 0 0 1 1.628 0l4.25 6.5a1.012 1.012 0 0 1 0 .639l-4.25 6.5a1.012 1.012 0 0 1-1.628 0l-4.25-6.5Z" /><path strokeLinecap="round" strokeLinejoin="round" d="M15 12a3 3 0 1 1-6 0 3 3 0 0 1 6 0Z" /></svg>;
const EyeSlashIcon = ({ className }) => <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}><path strokeLinecap="round" strokeLinejoin="round" d="M3.98 8.223A10.477 10.477 0 0 0 1.934 12C3.226 16.338 7.244 19.5 12 19.5c.993 0 1.953-.138 2.863-.395M6.228 6.228A10.451 10.451 0 0 1 12 4.5c4.756 0 8.773 3.162 10.065 7.498a10.522 10.522 0 0 1-4.293 5.774M6.228 6.228 3 3m3.228 3.228 1.5 1.5M21 21l-3-3m-3.228-3.228-1.5-1.5M12 15.75a3.75 3.75 0 0 1-3.75-3.75M9.345 12a2.25 2.25 0 0 0 2.25 2.25M15 12a2.25 2.25 0 0 0-2.25-2.25m-2.25 0a2.25 2.25 0 0 0-2.25 2.25m5.655 5.655L12 12" /></svg>;
const ChevronDown = ({ className }) => <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2} stroke="currentColor" className={className}><path strokeLinecap="round" strokeLinejoin="round" d="M19.5 8.25l-7.5 7.5-7.5-7.5" /></svg>;
const ChevronUp = ({ className }) => <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2} stroke="currentColor" className={className}><path strokeLinecap="round" strokeLinejoin="round" d="M4.5 15.75l7.5-7.5 7.5 7.5" /></svg>;

// Tab Icons
const PresentationChart = ({ className }) => <span className={className}>📊</span>;
const ArrowElbowRight = ({ className }) => <span className={className}>↪️</span>;
const ListNumbers = ({ className }) => <span className={className}>#️⃣</span>;
const Cube = ({ className }) => <span className={className}>📦</span>;
const ArrowsClockwise = ({ className }) => <span className={className}>🔄</span>;
const ChartLineUp = ({ className }) => <span className={className}>📈</span>;

// --- CONFIGURATION ---
const API_BASE = "http://localhost:4000";

const TABS = [
  { id: 'home', label: 'HOME' },
  { id: 'sycamore', label: 'SYCAMORE' },
  { id: 'client', label: 'SYCAMORE AND CLIENT' },
  { id: 'cft', label: 'CFT' },
  { id: 'weekly', label: 'WEEKLY UPDATE' },
  { id: 'tracker', label: 'TRACKER' },
  { id: 'project', label: 'PROJECT LIST' }
];

const USER_METRICS_MAP = [
  { label: 'Active', key: 'Number of Active Users' },
  { label: 'Full', key: 'Number of Full users' },
  { label: 'Read Only', key: 'Number of Read Only users' },
  { label: 'TLF', key: 'Number of TLF users' } 
];

// NEW MAPS FOR HOME DASHBOARD
const STAKEHOLDERS_MAP = [
  { label: 'CSM', key: 'CSM' },
  { label: 'Lead BA', key: 'Lead BA' },
  { label: 'Prod. Ops', key: 'Production Operation POC' },
  { label: 'Support', key: 'Support Lead' }
];

const VERSIONS_MAP = [
  { label: 'Product', key: 'Sycamore Informatics Product' },
  { label: 'Modules', key: 'Add-on Modules' }
];

const WEEKLY_ONLY_KEYS = [
    "Customer Sentiment Score",
    "Customer Sentiment Description"
];

const PRODUCT_TABS = [
  { key: 'currentState', title: 'Current State', icon: PresentationChart },
  { key: 'nextUp', title: 'Next Up', icon: ArrowElbowRight },
  { key: 'top3', title: 'Top 3 Items', icon: ListNumbers },
  { key: 'techStack', title: 'Tech Stack/Infra', icon: Wrench },
];

const CLIENT_SPECIFIC_TABS = [
  { key: 'deploymentDetails', title: 'Deployment Details', icon: Cube },
  { key: 'scheduledActivities', title: 'Scheduled Activities', icon: CalendarIcon },
  { key: 'productAlignment', title: 'Product Alignment', icon: ArrowsClockwise },
  { key: 'performanceMetrics', title: 'Performance Metrics', icon: ChartLineUp },
];

const SYCAMORE_SUBCATEGORIES = {
    "Customer & Engagement": ["Customer Since When / Initial Go Live", "Project Lifecycle Status", "Customer POC", "CSM", "Backup CSM", "Quality Lead", "Lead BA", "Production Operation POC", "Support Lead", "Technical Lead", "Specialist in Sycamore Informatics", "SME", "Support Team", "Escalation Matrix"],
    "Product & Versions": ["Sycamore Informatics Product", "Add-on Modules", "Next Planned Version in Development", "Next Planned Version of Release", "Major Features Used or Requested", "Release Notes"],
    "System & Infrastructure": ["Hosting Platform", "App Cloud", "Data Cloud", "Compute Cloud", "Database", "Architecture"],
    "Performance & Availability": ["CPU", "Memory", "System Availability", "RDP", "RTO", "RPO"],
    "Support & Operations": ["Critical Tickets", "High", "Medium", "Low", "Ticket Volume and Resolution Time", "Backlog tickets/issues", "Capacity Planned", "RTM", "RTMVE"],
    "Documents": ["SOW", "QM and Certification", "Product Documents", "Deployment Documents", "Consolidated Document"],
    "Licensing & Tools": ["Windows Licensing", "CAL", "SAL", "MS Office", "Other Licensing", "Adobe", "Notepad"],
    "Training & Onboarding": ["Training & Onboarding (Client Data)"],
};

const FULL_SUBCATEGORY_MAP = {
  "Customer Information": [
    { key: "Customer Name", category: "Client", initialValue: "" }, 
    { key: "Customer Location", category: "Client", initialValue: "" }, 
    { key: "Customer Description", category: "Client", initialValue: "" }
  ]
};

const getInitialFormState = () => {
  const state = {};
  Object.values(FULL_SUBCATEGORY_MAP).forEach(group => {
    group.forEach(item => {
      state[item.key] = item.initialValue;
    });
  });
  return state;
};

// --- HOOKS ---

// Debounce hook to delay search queries
function useDebounce(value, delay) {
  const [debouncedValue, setDebouncedValue] = useState(value);
  useEffect(() => {
    const handler = setTimeout(() => {
      setDebouncedValue(value);
    }, delay);

    return () => {
      clearTimeout(handler);
    };
  }, [value, delay]);
  return debouncedValue;
}

// --- HELPER COMPONENTS ---

// Highlight Component
const HighlightText = ({ text, query }) => {
    if (!query || !text) return text;
    const escapedQuery = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const parts = text.toString().split(new RegExp(`(${escapedQuery})`, 'gi'));
    return (
        <span>
            {parts.map((part, i) => 
                part.toLowerCase() === query.toLowerCase() ? (
                    <span key={i} className="bg-yellow-300 text-black font-bold rounded-sm px-0.5">{part}</span>
                ) : (part)
            )}
        </span>
    );
};

// Rich Text Editor
function RichTextEditor({ value, onChange, placeholder, readOnly = false }) {
    const editorRef = useRef(null);
    const [showColorPalette, setShowColorPalette] = useState(false);
  
    useEffect(() => {
      if (editorRef.current && value !== editorRef.current.innerHTML) {
        editorRef.current.innerHTML = value || "";
      }
    }, [value]);
  
    const handleCommand = (command, value = null) => {
      document.execCommand(command, false, value);
      editorRef.current.focus();
    };
  
    const changeFontSize = (direction) => {
      const currentSize = document.queryCommandValue('fontSize') || '3';
      let newSize = parseInt(currentSize, 10);
      newSize = direction === 'increase' ? Math.min(newSize + 1, 7) : Math.max(newSize - 1, 1);
      handleCommand('fontSize', newSize.toString());
    };
    
    const handleAddImage = () => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = 'image/*';
      input.onchange = e => {
        const file = e.target.files[0];
        if (file) {
          const reader = new FileReader();
          reader.onload = (event) => handleCommand('insertImage', event.target.result);
          reader.readAsDataURL(file);
        }
      };
      input.click();
    };
  
    const colors = ['#000000', '#16A34A', '#2563EB', '#DC2626', '#EAB308', '#9333EA', '#FFFFFF'];
  
    return (
      <div className="border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-800 text-black dark:text-white h-full flex flex-col">
        {!readOnly && (
          <div className="flex items-center flex-wrap gap-2 p-2 border-b border-gray-200 dark:border-gray-700 bg-gray-50 dark:bg-gray-700 rounded-t-lg shrink-0">
            <button type="button" onClick={() => handleCommand('bold')} className="px-2 py-1 rounded hover:bg-gray-200 dark:hover:bg-gray-600 font-bold">B</button>
            <button type="button" onClick={() => handleCommand('italic')} className="px-2 py-1 rounded hover:bg-gray-200 dark:hover:bg-gray-600 italic">I</button>
            <button type="button" onClick={() => handleCommand('underline')} className="px-2 py-1 rounded hover:bg-gray-200 dark:hover:bg-gray-600 underline">U</button>
            <button type="button" onClick={() => changeFontSize('increase')} className="px-2 py-1 rounded hover:bg-gray-200 dark:hover:bg-gray-600 text-lg">A+</button>
            <button type="button" onClick={() => changeFontSize('decrease')} className="px-2 py-1 rounded hover:bg-gray-200 dark:hover:bg-gray-600 text-xs">A-</button>
            <button type="button" onClick={handleAddImage} className="px-2 py-1 rounded hover:bg-gray-200 dark:hover:bg-gray-600">🖼️</button>
            <div className="relative">
              <button type="button" onClick={() => setShowColorPalette(!showColorPalette)} className="px-2 py-1 rounded hover:bg-gray-200 dark:hover:bg-gray-600">🎨</button>
              {showColorPalette && (
                <div className="absolute z-10 mt-2 w-32 bg-white dark:bg-gray-800 border rounded-md shadow-lg p-2 grid grid-cols-4 gap-1">
                  {colors.map(color => (
                    <button key={color} type="button" onClick={() => { handleCommand('foreColor', color); setShowColorPalette(false); }} className="w-6 h-6 rounded-full border" style={{ backgroundColor: color }}></button>
                  ))}
                </div>
              )}
            </div>
          </div>
        )}
        <div ref={editorRef} contentEditable={!readOnly} onInput={(e) => !readOnly && onChange(e.currentTarget.innerHTML)} className={`w-full p-4 focus:outline-none overflow-y-auto flex-1 min-h-[150px] ${readOnly ? 'cursor-default' : ''}`} placeholder={placeholder}></div>
      </div>
    );
}

// Sentiment Gauge (Larger & Editable Description)
// Improved SentimentGauge (Fixes "No Data" & Typing Issues)
function SentimentGauge({ score, description, isEditing, onScoreChange, onDescriptionChange }) {
    const [localDescription, setLocalDescription] = useState(description || "");
    const textareaRef = useRef(null);

    // FIX: Only depend on 'description' (prop). 
    // We only want to sync DOWN from the parent, not loop on our own state.
    useEffect(() => {
        if (document.activeElement !== textareaRef.current) {
            setLocalDescription(description || "");
        }
    }, [description]); 

    const handleChange = (e) => {
        const val = e.target.value;
        setLocalDescription(val);
        onDescriptionChange(val); 
    };

    const rotation = (score / 100) * 180 - 90;

    return (
        <div className="flex flex-col items-center justify-center p-4 bg-gray-50 dark:bg-gray-800 rounded-xl border border-gray-200 dark:border-gray-700 h-full shadow-sm w-full">
            <h3 className="text-sm font-bold uppercase text-gray-500 dark:text-gray-400 mb-4 tracking-wider">Customer Sentiment</h3>
            
            {/* GAUGE GRAPHIC */}
            <div className="relative w-64 h-32 mb-4">
                <svg viewBox="0 0 200 100" className="w-full h-full overflow-visible">
                    <defs>
                        <linearGradient id="sentiment-gradient" x1="0%" y1="0%" x2="100%" y2="0%">
                            <stop offset="0%" stopColor="#ef4444" />
                            <stop offset="50%" stopColor="#eab308" />
                            <stop offset="100%" stopColor="#22c55e" />
                        </linearGradient>
                    </defs>
                    <path d="M 10 100 A 90 90 0 0 1 190 100" fill="none" stroke="url(#sentiment-gradient)" strokeWidth="20" />
                    <g transform={`rotate(${rotation}, 100, 100)`} className="transition-transform duration-500 ease-out">
                        <path d="M 97 100 L 100 10 L 103 100 Z" fill="currentColor" className="text-gray-800 dark:text-white" />
                        <circle cx="100" cy="100" r="6" fill="currentColor" className="text-gray-800 dark:text-white" />
                    </g>
                </svg>
            </div>

            {/* EDITABLE SECTION */}
            <div className="text-center z-10 w-full px-2">
                {isEditing ? (
                    <>
                        <input type="range" min="0" max="100" value={score} onChange={(e) => onScoreChange(Number(e.target.value))} className="w-full h-2 bg-gray-200 rounded-lg appearance-none cursor-pointer dark:bg-gray-700 mb-4" />
                        <textarea 
                            ref={textareaRef}
                            value={localDescription}
                            onChange={handleChange}
                            placeholder="Sentiment details..."
                            maxLength={100}
                            className="w-full p-2 border rounded text-center text-sm bg-white dark:bg-gray-900 text-black dark:text-white h-20 resize-none focus:outline-none focus:ring-2 focus:ring-blue-500"
                        />
                        <div className="text-xs text-gray-400 mt-1 text-right w-full pr-1">
                            {localDescription.length}/100
                        </div>
                    </>
                ) : (
                    <p className="text-sm font-medium text-gray-600 dark:text-gray-300 italic min-h-[1.5rem] break-words">
                        {description || "No details available"}
                    </p>
                )}
            </div>
        </div>
    );
}

// Login Page
function LoginPage({ onLogin }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [showPassword, setShowPassword] = useState(false);

  const handleLogin = (e) => {
    e.preventDefault();

    if (username === 'admin' && password === password) {
      onLogin();
    } else {
      setError('Invalid username or password');
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-gray-100 dark:bg-gray-900 transition-colors duration-300">
      <div className="p-8 bg-white dark:bg-gray-800 rounded-xl shadow-2xl w-full max-w-sm m-4">
        <img src="/sycamore-logo.png" alt="Logo" className="h-12 mx-auto mb-6" />
        <h2 className="text-2xl font-bold text-center text-gray-900 dark:text-white mb-6">Login</h2>
        <form onSubmit={handleLogin}>
          <div className="mb-4">
            <label className="block text-gray-700 dark:text-gray-300 mb-2">Username</label>
            <input type="text" value={username} onChange={(e) => setUsername(e.target.value)} className="w-full px-4 py-2 border rounded-lg dark:bg-gray-700 dark:border-gray-600 dark:text-white" required />
          </div>
          <div className="mb-6">
            <label className="block text-gray-700 dark:text-gray-300 mb-2">Password</label>
            <div className="relative">
              <input type={showPassword ? 'text' : 'password'} value={password} onChange={(e) => setPassword(e.target.value)} className="w-full px-4 py-2 border rounded-lg dark:bg-gray-700 dark:border-gray-600 dark:text-white" required />
              <button type="button" onClick={() => setShowPassword(!showPassword)} className="absolute inset-y-0 right-0 px-3 flex items-center text-gray-500">
                {showPassword ? <EyeSlashIcon className="h-5 w-5" /> : <EyeIcon className="h-5 w-5" />}
              </button>
            </div>
          </div>
          {error && <p className="text-red-500 text-center text-sm mb-4">{error}</p>}
          <button type="submit" className="w-full bg-indigo-600 text-white py-2 rounded-lg hover:bg-indigo-700 transition-colors font-semibold">Login</button>
        </form>
      </div>
    </div>
  );
}

// Add Client Modal
function AddClientModal({ isVisible, onClose, onSubmit, initialData, isSubmitting }) {
    const [formData, setFormData] = useState(initialData);
    const [clientName, setClientName] = useState("");
    const [color, setColor] = useState("#4f46e5");
    const [logoFile, setLogoFile] = useState(null);

    useEffect(() => {
        if (isVisible) {
            setFormData(getInitialFormState());
            setClientName("");
            setColor("#4f46e5");
            setLogoFile(null);
        }
    }, [isVisible]);

    const handleChange = (key, value) => setFormData((prev) => ({ ...prev, [key]: value }));
    const handleFileChange = (e) => setLogoFile(e.target.files && e.target.files[0]);
    const handleSubmit = (e) => {
        e.preventDefault();
        const customerData = { Client: [], Sycamore: [], "Sycamore and Client": [] };
        Object.values(FULL_SUBCATEGORY_MAP).forEach(group => group.forEach(item => {
            const value = formData[item.key]?.trim() || '';
            if (value) customerData[item.category].push(`${item.key}: ${value}`);
        }));
        onSubmit({ customerName: clientName.trim(), customerData, logoFile, color });
    };

    if (!isVisible) return null;
    return (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
            <div className="bg-white dark:bg-gray-800 rounded-2xl shadow-2xl w-full max-w-5xl max-h-[90vh] overflow-y-auto">
                <form onSubmit={handleSubmit} className="text-gray-900 dark:text-white">
                    <div className="p-6 border-b border-gray-200 dark:border-gray-700 flex justify-between items-center sticky top-0 bg-white dark:bg-gray-800 z-10">
                        <h2 className="text-2xl font-bold">Add New Client</h2>
                        <button type="button" onClick={onClose} className="p-2 hover:bg-gray-100 dark:hover:bg-gray-700 rounded-full">✕</button>
                    </div>
                    <div className="p-6 space-y-6">
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                            <input type="text" placeholder="Client Name" value={clientName} onChange={e => setClientName(e.target.value)} className="p-4 text-lg border rounded-lg w-full text-black dark:text-white bg-white dark:bg-gray-900 border-gray-300 dark:border-gray-600" required />
                            <input type="file" onChange={handleFileChange} className="p-4 text-lg border rounded-lg w-full text-black dark:text-white bg-white dark:bg-gray-900 border-gray-300 dark:border-gray-600" />
                        </div>
                    </div>
                    <div className="p-6 border-t border-gray-200 dark:border-gray-700 flex justify-end gap-3 sticky bottom-0 bg-white dark:bg-gray-800">
                        <button type="button" onClick={onClose} className="px-6 py-3 text-lg rounded-lg border hover:bg-gray-50 dark:hover:bg-gray-700">Cancel</button>
                        <button type="submit" disabled={isSubmitting} className="px-6 py-3 text-lg bg-blue-600 text-white rounded-lg hover:bg-blue-700">{isSubmitting ? 'Saving...' : 'Save Client'}</button>
                    </div>
                </form>
            </div>
        </div>
    );
}

// Export Modal Component
function ExportModal({ isVisible, onClose, customerName, selectedWeek }) {
    const [isExporting, setIsExporting] = useState(false);
    const [exportData, setExportData] = useState(null);

    useEffect(() => {
        if (isVisible && customerName) {
            fetchExportData();
        }
    }, [isVisible, customerName, selectedWeek]);

    const fetchExportData = async () => {
        try {
            const res = await fetch(`${API_BASE}/api/customers/${encodeURIComponent(customerName)}/export?week=${selectedWeek}`);
            if (res.ok) {
                const data = await res.json();
                setExportData(data);
            }
        } catch (e) {
            console.error("Failed to fetch export data:", e);
        }
    };

    const exportToPDF = () => {
        if (!exportData) return;
        setIsExporting(true);
        try {
            const doc = new jsPDF();
            const pageWidth = doc.internal.pageSize.width;
            let y = 20;

            // Title
            doc.setFontSize(20);
            doc.setFont(undefined, 'bold');
            doc.text(`SCD Report: ${exportData.customerName}`, pageWidth / 2, y, { align: 'center' });
            y += 10;
            doc.setFontSize(10);
            doc.setFont(undefined, 'normal');
            doc.text(`Week: ${exportData.week} | Exported: ${new Date(exportData.exportedAt).toLocaleString()}`, pageWidth / 2, y, { align: 'center' });
            y += 15;

            // Helper function to add content with page breaks
            const addSection = (title, content, isTable = false) => {
                if (y > 250) { doc.addPage(); y = 20; }
                doc.setFontSize(14);
                doc.setFont(undefined, 'bold');
                doc.text(title, 14, y);
                y += 8;

                if (isTable && content && Array.isArray(content) && content.length > 0) {
                    const tableData = content.map(item => {
                        const [key, value] = item.split(/:(.*)/s);
                        return [key.trim(), value ? value.trim() : ''];
                    });
                    doc.autoTable({
                        startY: y,
                        head: [['Field', 'Value']],
                        body: tableData,
                        theme: 'striped',
                        headStyles: { fillColor: [79, 70, 229] },
                        styles: { fontSize: 9 },
                        margin: { left: 14, right: 14 }
                    });
                    y = doc.lastAutoTable.finalY + 10;
                } else if (!isTable && content) {
                    doc.setFontSize(10);
                    doc.setFont(undefined, 'normal');
                    if (Array.isArray(content)) {
                        content.forEach(item => {
                            if (y > 270) { doc.addPage(); y = 20; }
                            const [key, value] = item.split(/:(.*)/s);
                            doc.text(`• ${key.trim()}: ${value ? value.trim() : ''}`, 14, y);
                            y += 6;
                        });
                    } else if (typeof content === 'object') {
                        Object.entries(content).forEach(([key, val]) => {
                            if (y > 270) { doc.addPage(); y = 20; }
                            const displayVal = typeof val === 'object' ? JSON.stringify(val) : val;
                            doc.text(`• ${key}: ${displayVal || ''}`, 14, y);
                            y += 6;
                        });
                    } else {
                        if (y > 270) { doc.addPage(); y = 20; }
                        doc.text(content.toString(), 14, y);
                        y += 6;
                    }
                    y += 5;
                } else {
                    doc.setFontSize(10);
                    doc.setFont(undefined, 'italic');
                    doc.text('No data available', 14, y);
                    y += 8;
                }
            };

            // HOME Section
            addSection('HOME', null);
            if (exportData.home) {
                if (exportData.home.clientInfo && exportData.home.clientInfo.length > 0) {
                    addSection('Client Information', exportData.home.clientInfo, true);
                }
                if (exportData.home.stakeholders && exportData.home.stakeholders.length > 0) {
                    addSection('Stakeholders', exportData.home.stakeholders, true);
                }
                if (exportData.home.users && exportData.home.users.length > 0) {
                    addSection('Users', exportData.home.users, true);
                }
            }

            // SYCAMORE Section
            addSection('SYCAMORE', exportData.sycamore, true);

            // SYCAMORE AND CLIENT Section
            addSection('SYCAMORE AND CLIENT', exportData.sycamoreAndClient, true);

            // CFT Section
            addSection('CFT - Product Updates', null);
            if (exportData.cft && exportData.cft.productUpdates) {
                Object.entries(exportData.cft.productUpdates).forEach(([key, val]) => {
                    if (y > 270) { doc.addPage(); y = 20; }
                    doc.setFontSize(11);
                    doc.setFont(undefined, 'bold');
                    doc.text(key.replace(/([A-Z])/g, ' $1').trim(), 14, y);
                    y += 6;
                    doc.setFontSize(10);
                    doc.setFont(undefined, 'normal');
                    const cleanText = val ? val.replace(/<[^>]*>/g, '') : 'No data';
                    const splitText = doc.splitTextToSize(cleanText, 180);
                    splitText.forEach(line => {
                        if (y > 270) { doc.addPage(); y = 20; }
                        doc.text(line, 14, y);
                        y += 5;
                    });
                    y += 5;
                });
            }

            addSection('CFT - Client Specific Details', null);
            if (exportData.cft && exportData.cft.clientSpecificDetails) {
                Object.entries(exportData.cft.clientSpecificDetails).forEach(([key, val]) => {
                    if (y > 270) { doc.addPage(); y = 20; }
                    doc.setFontSize(11);
                    doc.setFont(undefined, 'bold');
                    doc.text(key.replace(/([A-Z])/g, ' $1').trim(), 14, y);
                    y += 6;
                    doc.setFontSize(10);
                    doc.setFont(undefined, 'normal');
                    const cleanText = val ? val.replace(/<[^>]*>/g, '') : 'No data';
                    const splitText = doc.splitTextToSize(cleanText, 180);
                    splitText.forEach(line => {
                        if (y > 270) { doc.addPage(); y = 20; }
                        doc.text(line, 14, y);
                        y += 5;
                    });
                    y += 5;
                });
            }

            // TRACKER Section
            addSection('TRACKER', null);
            if (exportData.tracker) {
                Object.entries(exportData.tracker).sort((a, b) => new Date(b[0]) - new Date(a[0])).forEach(([date, content]) => {
                    if (y > 260) { doc.addPage(); y = 20; }
                    doc.setFontSize(11);
                    doc.setFont(undefined, 'bold');
                    doc.text(date, 14, y);
                    y += 6;
                    doc.setFontSize(10);
                    doc.setFont(undefined, 'normal');
                    const cleanText = content ? content.replace(/<[^>]*>/g, '') : 'No content';
                    const splitText = doc.splitTextToSize(cleanText, 180);
                    splitText.forEach(line => {
                        if (y > 270) { doc.addPage(); y = 20; }
                        doc.text(line, 14, y);
                        y += 5;
                    });
                    y += 5;
                });
            }

            // PROJECT LIST Section
            addSection('PROJECT LIST', null);
            if (exportData.projectList) {
                Object.entries(exportData.projectList).forEach(([year, content]) => {
                    if (y > 260) { doc.addPage(); y = 20; }
                    doc.setFontSize(11);
                    doc.setFont(undefined, 'bold');
                    doc.text(`Year: ${year}`, 14, y);
                    y += 6;
                    doc.setFontSize(10);
                    doc.setFont(undefined, 'normal');
                    const cleanText = content ? content.replace(/<[^>]*>/g, '') : 'No content';
                    const splitText = doc.splitTextToSize(cleanText, 180);
                    splitText.forEach(line => {
                        if (y > 270) { doc.addPage(); y = 20; }
                        doc.text(line, 14, y);
                        y += 5;
                    });
                    y += 5;
                });
            }

            doc.save(`${exportData.customerName}_Report_${exportData.week}.pdf`);
        } catch (e) {
            console.error("PDF export error:", e);
            alert("Failed to export PDF");
        } finally {
            setIsExporting(false);
        }
    };

    const exportToCSV = () => {
        if (!exportData) return;
        setIsExporting(true);
        try {
            const rows = [['Category', 'Field', 'Value']];

            // Helper to add rows
            const addRows = (category, items) => {
                if (items && Array.isArray(items)) {
                    items.forEach(item => {
                        const [key, value] = item.split(/:(.*)/s);
                        rows.push([category, key.trim(), value ? value.trim().replace(/<[^>]*>/g, '') : '']);
                    });
                }
            };

            const addSectionRows = (category, content, isKeyValue = false) => {
                if (!content) return;
                if (isKeyValue && Array.isArray(content)) {
                    addRows(category, content);
                } else if (typeof content === 'object') {
                    Object.entries(content).forEach(([key, val]) => {
                        const displayVal = typeof val === 'object' ? JSON.stringify(val) : (val ? val.replace(/<[^>]*>/g, '') : '');
                        rows.push([category, key, displayVal]);
                    });
                }
            };

            // HOME
            if (exportData.home) {
                addRows('HOME - Client Info', exportData.home.clientInfo);
                addRows('HOME - Stakeholders', exportData.home.stakeholders);
                addRows('HOME - Users', exportData.home.users);
            }
            addRows('SYCAMORE', exportData.sycamore);
            addRows('SYCAMORE AND CLIENT', exportData.sycamoreAndClient);

            // CFT
            if (exportData.cft) {
                addSectionRows('CFT - Product Updates', exportData.cft.productUpdates);
                addSectionRows('CFT - Client Specific Details', exportData.cft.clientSpecificDetails);
            }

            // TRACKER
            if (exportData.tracker) {
                Object.entries(exportData.tracker).forEach(([date, content]) => {
                    rows.push(['TRACKER', date, content ? content.replace(/<[^>]*>/g, '') : '']);
                });
            }

            // PROJECT LIST
            if (exportData.projectList) {
                Object.entries(exportData.projectList).forEach(([year, content]) => {
                    rows.push(['PROJECT LIST', year, content ? content.replace(/<[^>]*>/g, '') : '']);
                });
            }

            const csvContent = rows.map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')).join('\n');
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            saveAs(blob, `${exportData.customerName}_Report_${exportData.week}.csv`);
        } catch (e) {
            console.error("CSV export error:", e);
            alert("Failed to export CSV");
        } finally {
            setIsExporting(false);
        }
    };

    const exportToExcel = () => {
        if (!exportData) return;
        setIsExporting(true);
        try {
            const wb = XLSX.utils.book_new();

            // Helper to create sheet from array
            const createSheetFromArray = (title, items) => {
                if (!items || !Array.isArray(items) || items.length === 0) return null;
                const data = [['Field', 'Value']];
                items.forEach(item => {
                    const [key, value] = item.split(/:(.*)/s);
                    data.push([key.trim(), value ? value.trim().replace(/<[^>]*>/g, '') : '']);
                });
                const ws = XLSX.utils.aoa_to_sheet(data);
                ws['!cols'] = [{ wch: 30 }, { wch: 60 }];
                return { title, ws };
            };

            // HOME sheets
            const homeInfo = createSheetFromArray('Home - Client Info', exportData.home?.clientInfo);
            const homeStake = createSheetFromArray('Home - Stakeholders', exportData.home?.stakeholders);
            const homeUsers = createSheetFromArray('Home - Users', exportData.home?.users);

            // SYCAMORE sheets
            const sycamoreSheet = createSheetFromArray('Sycamore', exportData.sycamore);
            const jointSheet = createSheetFromArray('Sycamore and Client', exportData.sycamoreAndClient);

            // Add sheets to workbook
            const addSheet = (sheetData) => {
                if (sheetData) {
                    XLSX.utils.book_append_sheet(wb, sheetData.ws, sheetData.title);
                }
            };

            addSheet(homeInfo);
            addSheet(homeStake);
            addSheet(homeUsers);
            addSheet(sycamoreSheet);
            addSheet(jointSheet);

            // CFT - Product Updates
            if (exportData.cft?.productUpdates) {
                const puData = [['Section', 'Content']];
                Object.entries(exportData.cft.productUpdates).forEach(([key, val]) => {
                    puData.push([key, val ? val.replace(/<[^>]*>/g, '') : '']);
                });
                const puWs = XLSX.utils.aoa_to_sheet(puData);
                puWs['!cols'] = [{ wch: 20 }, { wch: 80 }];
                XLSX.utils.book_append_sheet(wb, puWs, 'CFT - Product Updates');
            }

            // CFT - Client Specific Details
            if (exportData.cft?.clientSpecificDetails) {
                const csData = [['Section', 'Content']];
                Object.entries(exportData.cft.clientSpecificDetails).forEach(([key, val]) => {
                    csData.push([key, val ? val.replace(/<[^>]*>/g, '') : '']);
                });
                const csWs = XLSX.utils.aoa_to_sheet(csData);
                csWs['!cols'] = [{ wch: 25 }, { wch: 80 }];
                XLSX.utils.book_append_sheet(wb, csWs, 'CFT - Client Details');
            }

            // TRACKER
            if (exportData.tracker) {
                const trackerData = [['Date', 'Entry']];
                Object.entries(exportData.tracker).sort((a, b) => new Date(b[0]) - new Date(a[0])).forEach(([date, content]) => {
                    trackerData.push([date, content ? content.replace(/<[^>]*>/g, '') : '']);
                });
                const trackerWs = XLSX.utils.aoa_to_sheet(trackerData);
                trackerWs['!cols'] = [{ wch: 15 }, { wch: 80 }];
                XLSX.utils.book_append_sheet(wb, trackerWs, 'Tracker');
            }

            // PROJECT LIST
            if (exportData.projectList) {
                const plData = [['Year', 'Projects']];
                Object.entries(exportData.projectList).forEach(([year, content]) => {
                    plData.push([year, content ? content.replace(/<[^>]*>/g, '') : '']);
                });
                const plWs = XLSX.utils.aoa_to_sheet(plData);
                plWs['!cols'] = [{ wch: 10 }, { wch: 80 }];
                XLSX.utils.book_append_sheet(wb, plWs, 'Project List');
            }

            XLSX.writeFile(wb, `${exportData.customerName}_Report_${exportData.week}.xlsx`);
        } catch (e) {
            console.error("Excel export error:", e);
            alert("Failed to export Excel");
        } finally {
            setIsExporting(false);
        }
    };

    if (!isVisible) return null;

    return (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
            <div className="bg-white dark:bg-gray-800 rounded-2xl shadow-2xl w-full max-w-md">
                <div className="p-6 border-b border-gray-200 dark:border-gray-700 flex justify-between items-center">
                    <h2 className="text-2xl font-bold text-gray-900 dark:text-white">Export Client Data</h2>
                    <button onClick={onClose} className="p-2 hover:bg-gray-100 dark:hover:bg-gray-700 rounded-full text-gray-500">✕</button>
                </div>
                <div className="p-6 space-y-4">
                    <div className="text-center pb-4 border-b border-gray-200 dark:border-gray-700">
                        <p className="text-lg font-medium text-gray-900 dark:text-white">{customerName}</p>
                        <p className="text-sm text-gray-500">Week: {selectedWeek}</p>
                    </div>
                    <p className="text-sm text-gray-600 dark:text-gray-400">Select export format:</p>
                    <div className="grid grid-cols-1 gap-3">
                        <button
                            onClick={exportToPDF}
                            disabled={isExporting || !exportData}
                            className="flex items-center gap-3 p-4 border rounded-xl hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                        >
                            <FileDocumentIcon className="w-8 h-8 text-red-600" />
                            <div className="text-left">
                                <div className="font-bold text-gray-900 dark:text-white">PDF Document</div>
                                <div className="text-sm text-gray-500">Formatted report with tables</div>
                            </div>
                        </button>
                        <button
                            onClick={exportToExcel}
                            disabled={isExporting || !exportData}
                            className="flex items-center gap-3 p-4 border rounded-xl hover:bg-green-50 dark:hover:bg-green-900/20 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                        >
                            <TableIcon className="w-8 h-8 text-green-600" />
                            <div className="text-left">
                                <div className="font-bold text-gray-900 dark:text-white">Excel Spreadsheet</div>
                                <div className="text-sm text-gray-500">Multiple sheets organized data</div>
                            </div>
                        </button>
                        <button
                            onClick={exportToCSV}
                            disabled={isExporting || !exportData}
                            className="flex items-center gap-3 p-4 border rounded-xl hover:bg-blue-50 dark:hover:bg-blue-900/20 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                        >
                            <DownloadIcon className="w-8 h-8 text-blue-600" />
                            <div className="text-left">
                                <div className="font-bold text-gray-900 dark:text-white">CSV File</div>
                                <div className="text-sm text-gray-500">Comma-separated values</div>
                            </div>
                        </button>
                    </div>
                    {isExporting && <div className="text-center text-sm text-gray-500">Exporting...</div>}
                </div>
            </div>
        </div>
    );
}

// --- MAIN APP ---
export default function App() {
    // --- STATE ---
    const [isAuthenticated, setIsAuthenticated] = useState(false);
    const [customers, setCustomers] = useState([]);
    const [selectedCustomer, setSelectedCustomer] = useState("");
    const [selectedWeek, setSelectedWeek] = useState("");
    const [availableWeeks, setAvailableWeeks] = useState([]);
    const [query, setQuery] = useState("");
    const [data, setData] = useState({});
    const [masterData, setMasterData] = useState({}); 
    const [editingTab, setEditingTab] = useState(null); 
    const [editingFields, setEditingFields] = useState({});
    const [theme, setTheme] = useState('light');
    const [isLoading, setIsLoading] = useState(false);
    const [activeTab, setActiveTab] = useState("home");
    const [hiddenCustomers, setHiddenCustomers] = useState([]);
    const [showHidden, setShowHidden] = useState(false);
    
    // UI State for Home
    const [showDescription, setShowDescription] = useState(false);

    // Search Metrics
    const [searchMatches, setSearchMatches] = useState({});

    // Tabs
    const [productUpdateData, setProductUpdateData] = useState({});
    const [clientSpecificData, setClientSpecificData] = useState({});
    const [expandedCFT, setExpandedCFT] = useState({ product: true, client: true });

    // Weekly Update Tab State
    const [weeklyUpdateContent, setWeeklyUpdateContent] = useState("");
    const [isSavingWeeklyUpdate, setIsSavingWeeklyUpdate] = useState(false);

    // Tracker State
    const [trackerData, setTrackerData] = useState({});
    const [selectedTrackerDate, setSelectedTrackerDate] = useState(new Date().toISOString().split('T')[0]);
    const [trackerContent, setTrackerContent] = useState('');
    const [showAllTrackerEntries, setShowAllTrackerEntries] = useState(false); 
    const [isTrackerSaving, setIsTrackerSaving] = useState(false);

    // Project List State
    const [plData, setPLData] = useState({});
    const [selectedPLYear, setSelectedPLYear] = useState(new Date().getFullYear());
    const [plContent, setPLContent] = useState('');
    const [isPLSaving, setIsPLSaving] = useState(false);

    // Modal States
    const [isAddClientModalVisible, setIsAddClientModalVisible] = useState(false);
    const [isAddingClient, setIsAddingClient] = useState(false);
    const [isExportModalVisible, setIsExportModalVisible] = useState(false);
    const [openMenu, setOpenMenu] = useState(null);

    // Debounced search query
    const debouncedQuery = useDebounce(query, 300);
    const [sidebarCounts, setSidebarCounts] = useState({});

    // --- EFFECTS ---

    useEffect(() => {
        if (theme === 'dark') document.documentElement.classList.add('dark');
        else document.documentElement.classList.remove('dark');
    }, [theme]);

    useEffect(() => {
        const fetchInitial = async () => {
            try {
                const res = await fetch(`${API_BASE}/api/weeks`);
                const weeks = await res.json();
                setAvailableWeeks(weeks);
                if (weeks.length > 0) setSelectedWeek(weeks.find(w => w.isCurrent)?.value || weeks[0].value);
            } catch (e) { console.error(e); }
        };
        fetchInitial();
    }, []);

    useEffect(() => {
        if (!selectedWeek) return;
        const loadWeek = async () => {
            setIsLoading(true);
            try {
                const res = await fetch(`${API_BASE}/api/customers?week=${selectedWeek}`);
                const list = await res.json();
                setCustomers(list);
                if (list.length > 0 && !selectedCustomer) setSelectedCustomer(list[0]);
            } catch (e) { console.error(e); } finally { setIsLoading(false); }
        };
        loadWeek();
    }, [selectedWeek]);

    const fetchCustomerData = async () => {
        if (!selectedWeek || !selectedCustomer) return;
        
        setIsLoading(true);
        try {
            let url = `${API_BASE}/api/customers/${encodeURIComponent(selectedCustomer)}?week=${selectedWeek}`;
            const res = await fetch(url);
            const json = await res.json();

            if (debouncedQuery) {
                const searchRes = await fetch(`${API_BASE}/api/search?q=${encodeURIComponent(debouncedQuery)}&week=${selectedWeek}`);
                const searchJson = await searchRes.json();

                const newCounts = {};
                if (searchJson.results) {
                    Object.entries(searchJson.results).forEach(([custName, categories]) => {
                        let total = 0;
                        Object.values(categories).forEach(matches => {
                            if (Array.isArray(matches)) total += matches.length;
                        });
                        newCounts[custName] = total;
                    });
                }
                setSidebarCounts(newCounts);
            } else {
                setSidebarCounts({});
            }

            setData({ [selectedCustomer]: json });
            
            // FIX: Only update editingFields if we are NOT currently editing.
            // This prevents the UI from resetting while you are typing if a background fetch finishes.
            if (!editingTab) {
                setEditingFields(json);
            }

            if (selectedWeek !== 'master') {
                    const masterRes = await fetch(`${API_BASE}/api/customers/${encodeURIComponent(selectedCustomer)}?week=master`);
                    if (masterRes.ok) setMasterData(await masterRes.json());
            }

            fetch(`${API_BASE}/api/customers/${encodeURIComponent(selectedCustomer)}/product-update?week=${selectedWeek}`)
                .then(r => r.ok ? r.json() : { data: {} }).then(j => setProductUpdateData(j.data || {}));

            fetch(`${API_BASE}/api/customers/${encodeURIComponent(selectedCustomer)}/client-specific-details?week=${selectedWeek}`)
                .then(r => r.ok ? r.json() : { data: {} }).then(j => setClientSpecificData(j.data || {}));
            
            fetch(`${API_BASE}/api/customers/${encodeURIComponent(selectedCustomer)}/tracker`)
                    .then(r => r.ok ? r.json() : { data: {} }).then(j => setTrackerData(j.data || {}));

            fetch(`${API_BASE}/api/customers/${encodeURIComponent(selectedCustomer)}/project-list`)
                    .then(r => r.ok ? r.json() : { data: {} }).then(j => setPLData(j.data || {}));
            
        } catch (e) { console.error(e); } finally { setIsLoading(false); }
    };

    useEffect(() => {
        fetchCustomerData();
    }, [selectedCustomer, selectedWeek, debouncedQuery]);

    // Fetch Weekly Update when tab is active or week changes
    useEffect(() => {
        if (activeTab === 'weekly' && selectedWeek) {
            fetch(`${API_BASE}/api/weekly-update?week=${selectedWeek}`)
                .then(r => r.json())
                .then(data => setWeeklyUpdateContent(data.text || ""))
                .catch(err => console.error("Failed to load weekly update", err));
        }
    }, [activeTab, selectedWeek]);

    useEffect(() => {
        const dateKey = new Date(selectedTrackerDate).toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' });
        setTrackerContent(trackerData[dateKey] || '');
    }, [trackerData, selectedTrackerDate]);

    useEffect(() => {
        setPLContent(plData[selectedPLYear] || '');
    }, [plData, selectedPLYear]);

    useEffect(() => {
        setEditingTab(null);
    }, [activeTab, selectedCustomer]);

    // Search Metrics
    useEffect(() => {
        if (!query || !data[selectedCustomer]) {
            setSearchMatches({});
            return;
        }
        const q = query.toLowerCase();
        const countInObject = (obj) => {
            if (!obj) return 0;
            return JSON.stringify(obj).toLowerCase().split(q).length - 1;
        };
        const custData = data[selectedCustomer];
        setSearchMatches({
            home: countInObject(custData["Client"]),
            sycamore: countInObject(custData["Sycamore"]),
            client: countInObject(custData["Sycamore and Client"]),
            cft: countInObject(productUpdateData) + countInObject(clientSpecificData),
            tracker: countInObject(trackerData),
            project: countInObject(plData),
            weekly: countInObject(weeklyUpdateContent)
        });
    }, [query, data, productUpdateData, clientSpecificData, trackerData, plData, weeklyUpdateContent, selectedCustomer]);


    // --- HANDLERS ---

    const handleFieldChange = (category, itemIndex, eventOrValue) => {
        const value = typeof eventOrValue === 'object' && eventOrValue.target ? eventOrValue.target.value : eventOrValue;
        const customerData = data[selectedCustomer];
        
        if (!customerData || !customerData[category]) return;
        const originalItem = customerData[category][itemIndex] || "";
        const [subcategory] = originalItem.split(/:(.*)/s);
        const newItem = `${subcategory}: ${value}`;
        
        setEditingFields(prev => {
            const newCat = [...(prev[category] || customerData[category])];
            newCat[itemIndex] = newItem;
            return { ...prev, [category]: newCat };
        });
    };

    const updateFieldByKey = (category, key, newValue) => {
        const arr = (editingFields[category] || data[selectedCustomer][category] || []);
        const index = arr.findIndex(i => i.startsWith(key));
        if (index !== -1) {
            handleFieldChange(category, index, newValue);
        } else {
            setEditingFields(prev => {
                const existingCat = prev[category] || data[selectedCustomer][category] || [];
                const newCat = [...existingCat];
                newCat.push(`${key}: ${newValue}`);
                return { ...prev, [category]: newCat };
            });
        }
    };

    const handleSaveCurrentTab = async () => {
        setIsLoading(true);
        try {
            const promises = [];

            if (['home', 'sycamore', 'client'].includes(editingTab)) {
                const weeklyPayload = {};
                const masterPayload = {};

                for (const category in editingFields) {
                    if (category.startsWith('_')) continue;
                    weeklyPayload[category] = [];
                    masterPayload[category] = [];

                    editingFields[category].forEach(editedItem => {
                        const masterItems = masterData[category] || [];
                        const key = editedItem.split(':')[0];
                        const isMasterCandidate = masterItems.some(i => i.startsWith(key + ':'));

                        // If the key is in the explicit weekly list, send it to weekly.
                        // Otherwise, use the existing logic.
                        if (WEEKLY_ONLY_KEYS.includes(key) || !isMasterCandidate) weeklyPayload[category].push(editedItem);
                        else masterPayload[category].push(editedItem);
                    });
                }

                if (selectedWeek !== 'master') {
                    promises.push(fetch(`${API_BASE}/api/customers/${encodeURIComponent(selectedCustomer)}/data?week=${selectedWeek}`, { method: 'PUT', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(weeklyPayload) }));
                }
                promises.push(fetch(`${API_BASE}/api/customers/${encodeURIComponent(selectedCustomer)}/data?week=master`, { method: 'PUT', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(masterPayload) }));
            }

            if (editingTab === 'cft') {
                promises.push(fetch(`${API_BASE}/api/customers/${encodeURIComponent(selectedCustomer)}/product-update?week=${selectedWeek}`, { method: 'PUT', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(productUpdateData) }));
                promises.push(fetch(`${API_BASE}/api/customers/${encodeURIComponent(selectedCustomer)}/client-specific-details?week=${selectedWeek}`, { method: 'PUT', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(clientSpecificData) }));
            }

            if (editingTab === 'tracker') {
                 const dateKey = new Date(selectedTrackerDate).toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' });
                 promises.push(fetch(`${API_BASE}/api/customers/${encodeURIComponent(selectedCustomer)}/tracker`, {
                    method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ date: dateKey, content: trackerContent })
                }));
                setTrackerData(prev => ({ ...prev, [dateKey]: trackerContent }));
            }

            if (editingTab === 'project') {
                 promises.push(fetch(`${API_BASE}/api/customers/${encodeURIComponent(selectedCustomer)}/project-list`, {
                    method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ year: selectedPLYear, content: plContent })
                }));
                setPLData(prev => ({ ...prev, [selectedPLYear]: plContent }));
            }

            await Promise.all(promises);
            await fetch(`${API_BASE}/api/cache/clear`, { method: 'POST' });
            setEditingTab(null);
            fetchCustomerData(); 
        } catch (e) { alert("Error saving: " + e.message); } finally { setIsLoading(false); }
    };

    const handleSaveWeeklyUpdate = async () => {
        setIsSavingWeeklyUpdate(true);
        try {
            await fetch(`${API_BASE}/api/weekly-update?week=${selectedWeek}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ text: weeklyUpdateContent }),
            });
            alert("Weekly Update Saved!");
            fetchCustomerData(); // Re-fetch all data to update the cache
        } catch (e) { alert(e.message); } finally { setIsSavingWeeklyUpdate(false); }
    };

    const handleCancel = () => {
        setEditingTab(null);
        fetchCustomerData();
    };

    const handleAddClient = async (payload) => {
        setIsAddingClient(true);
        try {
            await fetch(`${API_BASE}/api/customers?week=${selectedWeek}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            fetchCustomerData();
            setIsAddClientModalVisible(false);
        } catch (e) { alert(e.message); } finally { setIsAddingClient(false); }
    };

    const toggleCustomerVisibility = (customerName) => {
        setHiddenCustomers(prev => prev.includes(customerName) ? prev.filter(c => c !== customerName) : [...prev, customerName]);
    };

    // --- RENDER HELPERS ---
    // Fix: Added extra "|| {}" at the end to prevent crashing if editingFields is null
    const currentData = (editingTab ? editingFields : (data[selectedCustomer] || {})) || {}; 
    const clientInfo = currentData["Client"] || [];
    const sycamoreInfo = currentData["Sycamore"] || [];
    const mixedInfo = currentData["Sycamore and Client"] || [];

    const getValue = (arr, key) => {
        if (!arr || !Array.isArray(arr)) return "";
        const item = arr.find(i => typeof i === 'string' && i.startsWith(key));
        return item ? item.split(/:(.*)/s)[1]?.trim() : "";
    };

    const renderField = (category, key) => {
        const arr = category === "Client" ? clientInfo : category === "Sycamore" ? sycamoreInfo : mixedInfo;
        const index = arr.findIndex(i => i.startsWith(key));
        const val = getValue(arr, key);

        if (editingTab && ['home', 'sycamore', 'client'].includes(activeTab) && index !== -1) {
            return (
                <input 
                    className="w-full bg-transparent border-b border-gray-400 dark:border-gray-500 px-1 py-0.5 text-base focus:outline-none focus:border-blue-500 text-black dark:text-white font-medium placeholder-gray-500"
                    value={val}
                    onChange={(e) => handleFieldChange(category, index, e)}
                />
            );
        }
        return (
            <span className="text-black dark:text-gray-100 font-medium break-words text-lg">
                <HighlightText text={val || "No Data"} query={query} />
            </span>
        );
    };

    const sidebarClass = theme === 'light' ? 'bg-blue-600 border-blue-700 text-white' : 'bg-gray-900 border-gray-800 text-gray-400';
    const mainClass = theme === 'light' ? 'bg-white text-black' : 'bg-black text-white';
    const headerClass = theme === 'light' ? 'bg-white border-gray-300' : 'bg-gray-900 border-gray-800';
    const cardClass = theme === 'light' ? 'bg-white border-gray-300 shadow-md' : 'bg-gray-800 border-gray-700';
    const textClass = theme === 'light' ? 'text-black' : 'text-white';
    const labelClass = theme === 'light' ? 'text-black font-bold' : 'text-gray-400 font-semibold';

    if (!isAuthenticated) return <div className={theme === 'dark' ? 'dark' : ''}><LoginPage onLogin={() => setIsAuthenticated(true)} /></div>;

    const groupedTrackerData = Object.keys(trackerData).sort((a,b) => new Date(b) - new Date(a)).reduce((acc, date) => {
        const [m, d, y] = date.split('/');
        const monthYear = new Date(y, parseInt(m, 10)-1).toLocaleString('default', { month: 'long', year: 'numeric' });
        if(!acc[monthYear]) acc[monthYear] = [];
        acc[monthYear].push(date);
        return acc;
    }, {});

    return (
        <div className={`flex h-screen overflow-hidden font-sans transition-colors duration-300 ${mainClass}`}>
            
            {/* SIDEBAR */}
            <aside className={`w-72 flex flex-col flex-shrink-0 z-20 h-full border-r ${sidebarClass}`}>
                <div className="p-6 h-24 flex items-center shrink-0">
                    <img src="/sycamore-logo.png" alt="SCD" className="h-10 w-auto mr-4 brightness-0 invert" />
                    <span className="font-black text-5xl tracking-tighter text-white">SCD</span>
                </div>
                <div className="flex-1 overflow-y-auto p-3 space-y-2 scrollbar-thin">
                    {customers
                        .filter(cust => showHidden || !hiddenCustomers.includes(cust))
                        .map(cust => {
                            const matchCount = sidebarCounts[cust] || 0;
                            const isHidden = hiddenCustomers.includes(cust);
                            return (
                                <div key={cust} className="relative group">
                                    <button
                                        onClick={() => { setSelectedCustomer(cust); setEditingTab(null); }}
                                        className={`w-full flex items-center justify-between gap-4 pl-4 pr-2 py-4 rounded-xl transition-all duration-200 text-left ${
                                            selectedCustomer === cust
                                                ? (theme === 'light' ? 'bg-white/20 text-white font-bold' : 'bg-gray-800 text-white border-l-4 border-blue-500')
                                                : `hover:bg-white/10 dark:hover:bg-gray-800 ${isHidden ? 'opacity-50' : ''}`
                                        }`}
                                    >
                                        <span className={`flex-1 text-base truncate uppercase tracking-wide ${isHidden ? 'line-through' : ''}`}>
                                            <HighlightText text={cust} query={debouncedQuery} />
                                        </span>
                                        {matchCount > 0 && (
                                            <span className="ml-2 bg-yellow-400 text-black text-xs font-bold px-2 py-0.5 rounded-full">
                                                {matchCount}
                                            </span>
                                        )}
                                        <button onClick={(e) => { e.stopPropagation(); setOpenMenu(openMenu === cust ? null : cust); }} className="p-2 rounded-full hover:bg-white/20 opacity-0 group-hover:opacity-100 transition-opacity">
                                            <EllipsisVerticalIcon className="w-5 h-5" />
                                        </button>
                                    </button>
                                    {openMenu === cust && (
                                        <div className="absolute right-4 top-12 mt-1 bg-white dark:bg-gray-700 rounded-md shadow-lg z-20 py-1 w-36">
                                            <button onClick={() => { setIsExportModalVisible(true); setOpenMenu(null); }} className="w-full text-left px-4 py-2 text-sm text-gray-800 dark:text-gray-200 hover:bg-gray-100 dark:hover:bg-gray-600 flex items-center gap-2">
                                                <DownloadIcon className="w-4 h-4" /> Export
                                            </button>
                                            <button onClick={() => { toggleCustomerVisibility(cust); setOpenMenu(null); }} className="w-full text-left px-4 py-2 text-sm text-gray-800 dark:text-gray-200 hover:bg-gray-100 dark:hover:bg-gray-600">
                                                {isHidden ? 'Unhide Client' : 'Hide Client'}
                                            </button>
                                        </div>
                                    )}
                                </div>
                            );
                        })}
                </div>
                <div className="p-5 border-t border-white/10 dark:border-gray-800">
                    <button onClick={() => setIsAddClientModalVisible(true)} className={`w-full flex items-center justify-center gap-3 py-3 rounded-lg transition shadow-md font-bold text-lg ${theme === 'light' ? 'bg-white text-blue-600 hover:bg-blue-50' : 'bg-blue-600 text-white hover:bg-blue-700'}`}>
                        <span>+</span> Add Client
                    </button>
                </div>
            </aside>

            {/* CONTENT AREA */}
            <div className={`flex-1 flex flex-col h-full relative overflow-hidden ${mainClass}`}>
                {/* HEADER */}
                <header className={`h-24 flex items-center justify-between px-8 border-b shrink-0 ${headerClass}`}>
                    <div className={`text-2xl font-bold uppercase tracking-widest ${textClass}`}>Sycamore Customer Dashboard</div>
                    <div className="flex items-center gap-6">
                        <select value={selectedWeek} onChange={(e) => setSelectedWeek(e.target.value)} className={`pl-4 pr-10 py-2 rounded-lg border bg-transparent text-base font-medium focus:ring-2 focus:ring-blue-500 cursor-pointer ${theme === 'light' ? 'border-gray-400 text-black' : 'border-gray-700 text-white bg-gray-900'}`}>
                            {availableWeeks.map(w => <option key={w.value} value={w.value}>{w.label}</option>)}
                        </select>
                        <div className="relative">
                            <input type="text" placeholder="Search..." value={query} onChange={(e) => setQuery(e.target.value)} className={`pl-4 pr-10 py-2 rounded-lg border bg-transparent text-base focus:ring-2 focus:ring-blue-500 w-64 ${theme === 'light' ? 'border-gray-400 text-black placeholder-gray-600' : 'border-gray-700 text-white placeholder-gray-500'}`} />
                            <span className="absolute right-3 top-2.5 text-gray-500 text-sm">🔍</span>
                        </div>
                        <button onClick={() => setShowHidden(!showHidden)} className={`p-2 rounded-lg ${theme === 'light' ? 'hover:bg-gray-200 text-blue-800' : 'hover:bg-gray-800 text-yellow-400'}`}>
                            {showHidden ? <EyeSlashIcon className="w-7 h-7" /> : <EyeIcon className="w-7 h-7" />}
                        </button>
                        <button onClick={() => setTheme(theme === 'light' ? 'dark' : 'light')} className={`p-2 rounded-lg text-2xl ${theme === 'light' ? 'hover:bg-gray-200 text-blue-800' : 'hover:bg-gray-800 text-yellow-400'}`}>
                            {theme === 'light' ? '🌙' : '☀️'}
                        </button>
                    </div>
                </header>

                {/* TABS */}
                <div className={`flex items-center border-b px-8 h-16 shrink-0 overflow-x-auto scrollbar-hide ${theme === 'light' ? 'bg-white border-gray-300' : 'bg-gray-900 border-gray-800'}`}>
                    {TABS.map(tab => {
                        const count = searchMatches[tab.id];
                        return (
                            <button
                                key={tab.id}
                                onClick={() => { setActiveTab(tab.id); if(editingTab !== tab.id) setEditingTab(null); }}
                                className={`px-8 h-full text-sm font-bold uppercase tracking-wider border-b-4 transition-colors whitespace-nowrap flex items-center gap-2 ${
                                    activeTab === tab.id 
                                    ? 'border-blue-600 text-blue-600 dark:border-blue-400 dark:text-blue-400' 
                                    : `border-transparent hover:text-gray-900 dark:hover:text-white ${theme === 'light' ? 'text-gray-600' : 'text-gray-400'}`
                                }`}
                            >
                                {tab.label}
                                {count > 0 && (
                                    <span className="bg-yellow-400 text-black text-xs px-2 py-0.5 rounded-full font-extrabold">{count}</span>
                                )}
                            </button>
                        )
                    })}
                    
                    {/* ACTIONS */}
                    {/* Only show Edit button if NOT in weekly update tab (since that has its own save logic) */}
                    {activeTab !== 'weekly' && selectedCustomer && (
                        <div className="ml-auto flex items-center gap-3">
                             {editingTab === activeTab ? (
                                <>
                                    <button onClick={handleSaveCurrentTab} className="bg-green-600 text-white px-4 py-2 text-sm rounded-lg uppercase font-bold hover:bg-green-700 shadow-sm">Save {activeTab}</button>
                                    <button onClick={handleCancel} className="bg-gray-500 text-white px-4 py-2 text-sm rounded-lg uppercase font-bold hover:bg-gray-600 shadow-sm">Cancel</button>
                                </>
                            ) : (
                                <button 
                                    onClick={() => { 
                                        // FIX: Safety check. If data is missing, default to empty object to prevent crash.
                                        const dataToEdit = data[selectedCustomer] || {};
                                        setEditingFields(dataToEdit); 
                                        setEditingTab(activeTab); 
                                    }} 
                                    className={`hover:text-blue-600 ${textClass} flex items-center gap-2 px-3 py-1 rounded hover:bg-black/5 dark:hover:bg-white/10`} 
                                    title={`Edit ${activeTab}`}>
                                    <PencilIcon className="w-5 h-5"/> <span className="text-sm font-bold uppercase">Edit {activeTab}</span>
                                </button>
                            )}
                        </div>
                    )}
                </div>

                {/* MAIN CONTENT */}
                <main className={`flex-1 overflow-y-auto p-10 scrollbar-thin ${mainClass}`}>
                    
                    {/* WEEKLY UPDATE TAB (Global) */}
                    {activeTab === 'weekly' ? (
                        <div className="max-w-5xl mx-auto h-full flex flex-col animate-fade-in">
                            <h2 className="text-3xl font-black mb-6 flex items-center gap-3 text-blue-600">
                                📝 Weekly Global Update
                            </h2>
                            <div className="flex-1 bg-white dark:bg-gray-800 rounded-xl shadow-xl overflow-hidden border dark:border-gray-700 flex flex-col">
                                <div className="flex-1 overflow-hidden p-6">
                                    <RichTextEditor 
                                        value={weeklyUpdateContent} 
                                        onChange={setWeeklyUpdateContent} 
                                        placeholder="Enter high-level weekly summary here..."
                                    />
                                </div>
                                <div className="p-4 border-t dark:border-gray-700 bg-gray-50 dark:bg-gray-900 flex justify-end">
                                    <button 
                                        onClick={handleSaveWeeklyUpdate}
                                        disabled={isSavingWeeklyUpdate}
                                        className="bg-blue-600 text-white px-8 py-3 rounded-lg font-bold hover:bg-blue-700 shadow-lg"
                                    >
                                        {isSavingWeeklyUpdate ? 'Saving...' : 'Save Weekly Update'}
                                    </button>
                                </div>
                            </div>
                        </div>
                    ) : selectedCustomer ? (
                        <>
                            {/* HOME TAB (REDESIGNED) */}
                            {activeTab === 'home' && (
                                <div className="w-full h-full flex flex-col space-y-6 animate-fade-in">
                                    
                                    {/* 1. Header Area: Logo, Name, Location & Dropdown Description */}
                                    <div className="flex flex-col gap-6">
                                        <div className="flex gap-6 items-center">
                                            <div className={`w-28 h-28 border rounded-2xl flex items-center justify-center p-4 shrink-0 shadow-sm ${theme === 'light' ? 'bg-white border-gray-300' : 'bg-white border-gray-700'}`}>
                                                <img src={`/${selectedCustomer}.png`} alt="logo" className="max-w-full max-h-full object-contain" />
                                            </div>
                                            <div className="space-y-2">
                                                <h1 className={`text-4xl font-black uppercase tracking-tighter ${textClass}`}>
                                                    <HighlightText text={selectedCustomer} query={query} />
                                                </h1>
                                                <div className="flex items-center gap-2">
                                                    <div className={`text-lg flex items-center gap-2 font-medium ${theme === 'light' ? 'text-black' : 'text-gray-300'}`}>
                                                        <span>📍</span> {renderField("Client", "Customer Location")}
                                                    </div>
                                                    {/* Dropdown Toggle */}
                                                    <button onClick={() => setShowDescription(!showDescription)} className="p-1 rounded-full hover:bg-black/10 dark:hover:bg-white/10 transition-colors">
                                                        <ChevronDown className={`w-5 h-5 transition-transform duration-200 ${showDescription ? 'rotate-180' : ''}`} />
                                                    </button>
                                                </div>
                                                
                                                {/* Collapsible Description */}
                                                {showDescription && (
                                                    <div className={`text-base italic mt-2 max-w-4xl leading-relaxed p-4 border-l-4 rounded-r-md animate-fade-in ${theme === 'light' ? 'text-black bg-gray-50 border-gray-300' : 'text-gray-400 bg-gray-800 border-gray-600'}`}>
                                                        {renderField("Client", "Customer Description")}
                                                    </div>
                                                )}
                                            </div>
                                        </div>
                                    </div>

                                    {/* 2. Four-Column Grid Layout */}
                                    <div className="grid grid-cols-1 xl:grid-cols-4 gap-6 w-full flex-1">
                                        
                                        {/* Column 1: Customer Sentiment (Moved from Right) */}
                                        <div className={`flex flex-col items-center ${cardClass} rounded-xl p-4 h-full`}>
                                             <SentimentGauge
                                                score={Number(getValue(clientInfo, "Customer Sentiment Score")) || 50}
                                                description={getValue(clientInfo, "Customer Sentiment Description")}
                                                isEditing={editingTab === 'home'}
                                                onScoreChange={(val) => updateFieldByKey("Client", "Customer Sentiment Score", val)}
                                                onDescriptionChange={(val) => updateFieldByKey("Client", "Customer Sentiment Description", val)}
                                            />
                                        </div>

                                        {/* Column 2: Users Table (Moved from Left & Compacted) */}
                                        <div className={`flex flex-col ${cardClass} rounded-xl p-6 h-full overflow-hidden`}>
                                            <h3 className={`text-sm font-bold uppercase mb-4 tracking-wider text-center ${labelClass}`}>Users</h3>
                                            <div className="flex-1 overflow-auto">
                                                <table className="w-full text-sm border-collapse">
                                                    <tbody>
                                                        {USER_METRICS_MAP.map((metric) => (
                                                            <tr key={metric.label}>
                                                                <td className={`font-bold p-3 border-b border-r ${theme === 'light' ? 'bg-gray-50 text-gray-700 border-gray-200' : 'bg-gray-900 text-gray-300 border-gray-700'}`}>{metric.label}</td>
                                                                <td className={`p-3 border-b text-center font-mono ${theme === 'light' ? 'text-black border-gray-200' : 'text-white border-gray-700'}`}>
                                                                    {renderField("Client", metric.key)}
                                                                </td>
                                                            </tr>
                                                        ))}
                                                    </tbody>
                                                </table>
                                            </div>
                                        </div>

                                        {/* Column 3: Key Stakeholders (New Table from Sycamore Data) */}
                                        <div className={`flex flex-col ${cardClass} rounded-xl p-6 h-full overflow-hidden`}>
                                            <h3 className={`text-sm font-bold uppercase mb-4 tracking-wider text-center ${labelClass}`}>Key Stakeholders</h3>
                                            <div className="flex-1 overflow-auto">
                                                <table className="w-full text-sm border-collapse">
                                                    <tbody>
                                                        {STAKEHOLDERS_MAP.map((item) => (
                                                            <tr key={item.label}>
                                                                <td className={`font-bold p-3 border-b border-r w-1/3 ${theme === 'light' ? 'bg-gray-50 text-gray-700 border-gray-200' : 'bg-gray-900 text-gray-300 border-gray-700'}`}>{item.label}</td>
                                                                <td className={`p-3 border-b ${theme === 'light' ? 'text-black border-gray-200' : 'text-white border-gray-700'}`}>
                                                                    {/* Note: We pull from "Sycamore" category even though we are on Home tab */}
                                                                    {renderField("Sycamore", item.key)}
                                                                </td>
                                                            </tr>
                                                        ))}
                                                    </tbody>
                                                </table>
                                            </div>
                                        </div>

                                        {/* Column 4: Versions (New Table from Sycamore Data) */}
                                        <div className={`flex flex-col ${cardClass} rounded-xl p-6 h-full overflow-hidden`}>
                                            <h3 className={`text-sm font-bold uppercase mb-4 tracking-wider text-center ${labelClass}`}>Versions & Product</h3>
                                            <div className="flex-1 overflow-auto">
                                                <table className="w-full text-sm border-collapse">
                                                    <tbody>
                                                        {VERSIONS_MAP.map((item) => (
                                                            <tr key={item.label}>
                                                                <td className={`font-bold p-3 border-b border-r w-1/3 ${theme === 'light' ? 'bg-gray-50 text-gray-700 border-gray-200' : 'bg-gray-900 text-gray-300 border-gray-700'}`}>{item.label}</td>
                                                                <td className={`p-3 border-b ${theme === 'light' ? 'text-black border-gray-200' : 'text-white border-gray-700'}`}>
                                                                    {renderField("Sycamore", item.key)}
                                                                </td>
                                                            </tr>
                                                        ))}
                                                    </tbody>
                                                </table>
                                            </div>
                                        </div>

                                    </div>
                                </div>
                            )}

                            {/* SYCAMORE TAB */}
                            {activeTab === 'sycamore' && (
                                <div className="grid grid-cols-1 md:grid-cols-2 gap-8 animate-fade-in">
                                    {Object.entries(SYCAMORE_SUBCATEGORIES).map(([title, keys]) => (
                                        <div key={title} className={`border rounded-2xl p-6 ${cardClass}`}>
                                            <h3 className="font-bold text-blue-600 dark:text-blue-400 mb-6 border-b border-gray-300 dark:border-gray-700 pb-3 flex items-center gap-3 text-lg">
                                                <Wrench className="w-5 h-5"/> {title.toUpperCase()}
                                            </h3>
                                            <div className="space-y-4">
                                                {keys.map(key => (
                                                    <div key={key} className="flex flex-col">
                                                        <span className={`text-sm uppercase font-bold mb-1 ${theme === 'light' ? 'text-black' : 'text-gray-500'}`}>{key}</span>
                                                        <div className="text-lg">{renderField("Sycamore", key)}</div>
                                                    </div>
                                                ))}
                                            </div>
                                        </div>
                                    ))}
                                </div>
                            )}

                            {/* JOINT INFO TAB */}
                            {activeTab === 'client' && (
                                <div className={`p-8 border rounded-2xl animate-fade-in ${cardClass}`}>
                                    <h3 className="font-bold text-purple-600 dark:text-purple-400 mb-8 text-xl border-b border-gray-300 dark:border-gray-700 pb-3">JOINT INFORMATION</h3>
                                    <div className="grid grid-cols-1 md:grid-cols-2 gap-10">
                                        {mixedInfo.map((item, idx) => {
                                            const [key] = item.split(/:(.*)/s);
                                            return (
                                                <div key={idx} className="flex flex-col">
                                                    <span className={`text-sm uppercase font-bold mb-1 ${theme === 'light' ? 'text-black' : 'text-gray-500'}`}>{key}</span>
                                                    {renderField("Sycamore and Client", key)}
                                                </div>
                                            )
                                        })}
                                    </div>
                                </div>
                            )}

                            {/* CFT TAB */}
                            {activeTab === 'cft' && (
                                <div className="space-y-8 animate-fade-in max-w-[1600px] mx-auto">
                                    <div className={`border rounded-2xl overflow-hidden ${cardClass}`}>
                                        <button onClick={() => setExpandedCFT(p => ({...p, product: !p.product}))} className={`w-full flex items-center justify-between p-6 bg-blue-50 dark:bg-blue-900/20 hover:bg-blue-100 dark:hover:bg-blue-900/30 transition-colors border-b ${theme === 'light' ? 'border-gray-200' : 'border-gray-700'}`}>
                                            <h3 className="text-xl font-bold text-blue-600 dark:text-blue-400 flex items-center gap-3"><PresentationChart className="w-6 h-6"/> Product Updates</h3>
                                            {expandedCFT.product ? <ChevronUp className="w-6 h-6 text-blue-500"/> : <ChevronDown className="w-6 h-6 text-blue-500"/>}
                                        </button>
                                        {expandedCFT.product && (
                                            <div className="p-6 grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-6 bg-white dark:bg-gray-800">
                                                {PRODUCT_TABS.map(tab => {
                                                    const Icon = tab.icon;
                                                    return (
                                                        <div key={tab.key} className="flex flex-col h-full border rounded-xl overflow-hidden dark:border-gray-700 shadow-sm">
                                                            <div className="bg-gray-50 dark:bg-gray-700/50 p-3 border-b dark:border-gray-700 flex items-center gap-2 font-bold text-sm text-gray-700 dark:text-gray-200"><Icon className="w-5 h-5 text-blue-500"/> {tab.title}</div>
                                                            <div className="flex-1 bg-white dark:bg-gray-800"><RichTextEditor value={productUpdateData[tab.key] || ''} readOnly={editingTab !== 'cft'} onChange={(html) => setProductUpdateData(prev => ({...prev, [tab.key]: html}))} placeholder={`...`} /></div>
                                                        </div>
                                                    );
                                                })}
                                            </div>
                                        )}
                                    </div>
                                    <div className={`border rounded-2xl overflow-hidden ${cardClass}`}>
                                        <button onClick={() => setExpandedCFT(p => ({...p, client: !p.client}))} className={`w-full flex items-center justify-between p-6 bg-purple-50 dark:bg-purple-900/20 hover:bg-purple-100 dark:hover:bg-purple-900/30 transition-colors border-b ${theme === 'light' ? 'border-gray-200' : 'border-gray-700'}`}>
                                            <h3 className="text-xl font-bold text-purple-600 dark:text-purple-400 flex items-center gap-3"><Cube className="w-6 h-6"/> Client Specific Details</h3>
                                            {expandedCFT.client ? <ChevronUp className="w-6 h-6 text-purple-500"/> : <ChevronDown className="w-6 h-6 text-purple-500"/>}
                                        </button>
                                        {expandedCFT.client && (
                                            <div className="p-6 grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-6 bg-white dark:bg-gray-800">
                                                {CLIENT_SPECIFIC_TABS.map(tab => {
                                                    const Icon = tab.icon;
                                                    return (
                                                        <div key={tab.key} className="flex flex-col h-full border rounded-xl overflow-hidden dark:border-gray-700 shadow-sm">
                                                            <div className="bg-gray-50 dark:bg-gray-700/50 p-3 border-b dark:border-gray-700 flex items-center gap-2 font-bold text-sm text-gray-700 dark:text-gray-200"><Icon className="w-5 h-5 text-purple-500"/> {tab.title}</div>
                                                            <div className="flex-1 bg-white dark:bg-gray-800"><RichTextEditor value={clientSpecificData[tab.key] || ''} readOnly={editingTab !== 'cft'} onChange={(html) => setClientSpecificData(prev => ({...prev, [tab.key]: html}))} placeholder={`...`} /></div>
                                                        </div>
                                                    );
                                                })}
                                            </div>
                                        )}
                                    </div>
                                </div>
                            )}

                            {/* TRACKER TAB */}
                            {activeTab === 'tracker' && (
                                <div className={`flex flex-col h-full rounded-2xl shadow-xl overflow-hidden animate-fade-in ${cardClass}`}>
                                    <div className="p-6 border-b flex justify-between items-center bg-gray-50 dark:bg-gray-800 dark:border-gray-700">
                                        <h2 className="text-2xl font-bold flex items-center gap-3 text-indigo-600 dark:text-indigo-400">
                                            <CalendarIcon className="w-7 h-7"/> Tracker
                                        </h2>
                                        <div className="flex gap-4">
                                            <button onClick={() => setShowAllTrackerEntries(!showAllTrackerEntries)} className={`px-4 py-2 rounded-lg font-bold transition ${showAllTrackerEntries ? 'bg-indigo-600 text-white' : 'bg-white dark:bg-gray-700 border hover:bg-gray-50'}`}>
                                                {showAllTrackerEntries ? 'Back to Editor' : 'View All Entries'}
                                            </button>
                                            {!showAllTrackerEntries && (
                                                <input type="date" value={selectedTrackerDate} onChange={e => setSelectedTrackerDate(e.target.value)} className="p-2 border rounded-lg text-black dark:text-white bg-white dark:bg-gray-900 dark:border-gray-600" disabled={editingTab === 'tracker'} />
                                            )}
                                        </div>
                                    </div>
                                    <div className="flex-1 p-6 overflow-hidden flex flex-col">
                                        {showAllTrackerEntries ? (
                                            <div className="overflow-y-auto pr-4 space-y-8 h-full">
                                                {Object.keys(groupedTrackerData).map(monthYear => (
                                                    <div key={monthYear} className="space-y-4">
                                                        <h3 className="sticky top-0 bg-white dark:bg-gray-900 z-10 py-2 text-xl font-bold text-gray-500 border-b">{monthYear}</h3>
                                                        {groupedTrackerData[monthYear].map(date => (
                                                            <div key={date} className="border p-4 rounded-xl dark:border-gray-700">
                                                                <div className="font-bold mb-2 text-indigo-600">{date}</div>
                                                                <div className="prose dark:prose-invert max-w-none" dangerouslySetInnerHTML={{__html: trackerData[date]}} />
                                                            </div>
                                                        ))}
                                                    </div>
                                                ))}
                                                {Object.keys(groupedTrackerData).length === 0 && <p className="text-center italic text-gray-500">No entries found.</p>}
                                            </div>
                                        ) : (
                                            <>
                                                <div className="mb-2 text-sm text-gray-500 font-bold">Entry for: {new Date(selectedTrackerDate).toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })}</div>
                                                <div className="flex-1 border rounded-xl overflow-hidden dark:border-gray-700">
                                                    <RichTextEditor value={trackerContent} onChange={setTrackerContent} readOnly={editingTab !== 'tracker'} placeholder="No entry for this date." />
                                                </div>
                                            </>
                                        )}
                                    </div>
                                </div>
                            )}

                            {/* PROJECT LIST TAB */}
                            {activeTab === 'project' && (
                                <div className={`flex flex-col h-full rounded-2xl shadow-xl overflow-hidden animate-fade-in ${cardClass}`}>
                                    <div className="p-6 border-b flex justify-between items-center bg-gray-50 dark:bg-gray-800 dark:border-gray-700">
                                        <h2 className="text-2xl font-bold flex items-center gap-3 text-yellow-600 dark:text-yellow-400">
                                            <ProjectListIcon className="w-7 h-7"/> Project List
                                        </h2>
                                        <select value={selectedPLYear} onChange={(e) => setSelectedPLYear(Number(e.target.value))} className="p-2 border rounded-lg text-black dark:text-white bg-white dark:bg-gray-900 dark:border-gray-600" disabled={editingTab === 'project'}>
                                            {[2024, 2025, 2026, 2027].map(year => <option key={year} value={year}>{year}</option>)}
                                        </select>
                                    </div>
                                    <div className="flex-1 p-6 overflow-hidden flex flex-col">
                                         <div className="mb-2 text-sm text-gray-500 font-bold">List for Year: {selectedPLYear}</div>
                                         <div className="flex-1 border rounded-xl overflow-hidden dark:border-gray-700">
                                            <RichTextEditor value={plContent} onChange={setPLContent} readOnly={editingTab !== 'project'} placeholder={`No project list found for ${selectedPLYear}.`} />
                                        </div>
                                    </div>
                                </div>
                            )}

                        </>
                    ) : (
                        <div className="flex flex-col items-center justify-center h-full text-gray-400">
                            <span className="text-8xl mb-6">👈</span>
                            <p className="text-2xl font-medium">Select a Client from the sidebar</p>
                        </div>
                    )}
                </main>
            </div>

            {/* MODALS */}
            <AddClientModal isVisible={isAddClientModalVisible} onClose={() => setIsAddClientModalVisible(false)} onSubmit={handleAddClient} initialData={{}} isSubmitting={isAddingClient} />
            <ExportModal 
                isVisible={isExportModalVisible} 
                onClose={() => setIsExportModalVisible(false)} 
                customerName={selectedCustomer} 
                selectedWeek={selectedWeek} 
            />
        </div>
    );
}
