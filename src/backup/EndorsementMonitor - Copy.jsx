import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, Legend, ResponsiveContainer, LabelList } from 'recharts';
import { Hash, DollarSign, RefreshCw, AlertCircle, HelpCircle } from 'lucide-react';

// ============================================
// UTILITY FUNCTIONS
// ============================================

const parseNumber = (value) => {
  if (value === null || value === undefined) return 0;
  const cleanValue = String(value).replace(/,/g, '');
  const parsed = parseFloat(cleanValue);
  return isNaN(parsed) ? 0 : parsed;
};

const formatNumber = (num) => {
  return new Intl.NumberFormat('en-US', { 
    minimumFractionDigits: 2, 
    maximumFractionDigits: 2 
  }).format(num);
};

const formatPercent = (num) => `${formatNumber(Math.abs(num))}%`;

const calculateNetFlow = (endorsements, pullouts) => {
  if (pullouts === 0 && endorsements > 0) return { value: -1, isSpecial: true };
  if (pullouts === 0 && endorsements === 0) return { value: 0, isSpecial: true };
  return { value: endorsements / pullouts, isSpecial: false };
};

const formatNetFlow = (netFlowObj) => netFlowObj.isSpecial ? '−' : formatNumber(netFlowObj.value);

const getNetFlowIndicator = (netFlowObj) => {
  if (netFlowObj.isSpecial) return { icon: '−', color: 'text-gray-400' };
  if (netFlowObj.value > 1) return { icon: '↑', color: 'text-green-600' };
  return { icon: '↓', color: 'text-red-600' };
};

const getGrowthIndicator = (value) => {
  if (value > 0) return { icon: '↑', color: 'text-green-600' };
  if (value < 0) return { icon: '↓', color: 'text-red-600' };
  return { icon: '−', color: 'text-gray-400' };
};

const normalizeDateForComparison = (dateStr) => {
  const date = new Date(dateStr);
  date.setHours(0, 0, 0, 0);
  return date;
};

const validateSheetStructure = (workbook) => {
  const errors = [];
  const requiredSheets = ['DAILY', 'BOM', 'CAMPAIGN', 'CAMPAIGN_BOM'];
  requiredSheets.forEach(sheet => {
    if (!workbook.SheetNames.includes(sheet)) {
      errors.push(`Missing required sheet: "${sheet}"`);
    }
  });
  return errors;
};

const InfoTooltip = ({ text }) => {
  const [show, setShow] = useState(false);
  return (
    <div className="relative inline-block">
      <HelpCircle 
        className="h-4 w-4 text-gray-400 cursor-help" 
        onMouseEnter={() => setShow(true)}
        onMouseLeave={() => setShow(false)}
      />
      {show && (
        <div className="absolute z-10 w-64 p-2 text-xs bg-gray-800 text-white rounded shadow-lg -top-2 left-6 whitespace-normal">
          {text}
        </div>
      )}
    </div>
  );
};

const MultiSelectDropdown = ({ label, options, value, onChange }) => {
  const [isOpen, setIsOpen] = useState(false);
  const dropdownRef = React.useRef(null);

  React.useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleToggle = (option) => {
    if (value.includes(option)) {
      onChange(value.filter(v => v !== option));
    } else {
      onChange([...value, option]);
    }
  };

  return (
    <div className="relative" ref={dropdownRef}>
      <label className="block text-xs font-medium text-gray-600 mb-2">{label}</label>
      <button
        type="button"
        onClick={() => setIsOpen(!isOpen)}
        className="w-full p-2 border border-gray-300 rounded text-sm text-left bg-white hover:border-gray-400 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 flex items-center justify-between"
      >
        <span className={value.length === 0 ? 'text-gray-400' : 'text-gray-900'}>
          {value.length === 0 ? 'Select months...' : value.join(', ')}
        </span>
        <svg className={`w-5 h-5 text-gray-400 transition-transform ${isOpen ? 'rotate-180' : ''}`} fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
        </svg>
      </button>
      {isOpen && (
        <div className="absolute z-50 w-full mt-1 bg-white border border-gray-300 rounded shadow-lg max-h-60 overflow-y-auto">
          {options.map((option) => (
            <div
              key={option}
              onClick={() => handleToggle(option)}
              className="flex items-center px-3 py-2 hover:bg-gray-50 cursor-pointer"
            >
              <div className={`w-5 h-5 mr-3 border-2 rounded flex items-center justify-center ${
                value.includes(option) 
                  ? 'bg-indigo-600 border-indigo-600' 
                  : 'border-gray-300 bg-white'
              }`}>
                {value.includes(option) && (
                  <svg className="w-3 h-3 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={3} d="M5 13l4 4L19 7" />
                  </svg>
                )}
              </div>
              <span className="text-sm text-gray-700">{option}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default function EndorsementMonitor() {
  const [activeTab, setActiveTab] = useState('overall');
  const [dailyData, setDailyData] = useState([]);
  const [bomData, setBomData] = useState([]);
  const [campaignData, setCampaignData] = useState([]);
  const [campaignBomData, setCampaignBomData] = useState([]);
  const [fieldDailyData, setFieldDailyData] = useState([]);
  const [fieldBomData, setFieldBomData] = useState([]);
  const [fieldCampaignData, setFieldCampaignData] = useState([]);
  const [perAreaData, setPerAreaData] = useState([]);
  const [selectedMonth, setSelectedMonth] = useState('');
  const [selectedProductType, setSelectedProductType] = useState('');
  const [selectedClient, setSelectedClient] = useState('');
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [viewMode, setViewMode] = useState('count');
  const [selectedMonths, setSelectedMonths] = useState([]);
  const [selectedComparisonProductType, setSelectedComparisonProductType] = useState('');
  const [clientRankingView, setClientRankingView] = useState('all');
  const [fieldStartDate, setFieldStartDate] = useState('');
  const [fieldEndDate, setFieldEndDate] = useState('');
  const [availableMonths, setAvailableMonths] = useState([]);
  const [availableProductTypes, setAvailableProductTypes] = useState([]);
  const [availableClients, setAvailableClients] = useState([]);
  const [availableAreas, setAvailableAreas] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  const handleFileUpload = async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    setLoading(true);
    setError(null);
    
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });

      const validationErrors = validateSheetStructure(workbook);
      if (validationErrors.length > 0) {
        throw new Error(validationErrors.join('\n'));
      }

      let dailyJson = [];
      let bomJson = [];
      let campaignJson = [];
      let campaignBomJson = [];
      let fieldDailyJson = [];
      let fieldBomJson = [];
      let fieldCampaignJson = [];
      let perAreaJson = [];

      if (workbook.SheetNames.includes('DAILY')) {
        const dailySheet = workbook.Sheets['DAILY'];
        dailyJson = XLSX.utils.sheet_to_json(dailySheet, { raw: false });
        dailyJson = dailyJson.map(row => {
          if (row.DATE) {
            const date = new Date(row.DATE);
            const monthName = date.toLocaleString('en-US', { month: 'long' }).toUpperCase();
            return { ...row, MONTH: monthName };
          }
          return row;
        });
        setDailyData(dailyJson);
      }

      if (workbook.SheetNames.includes('BOM')) {
        const bomSheet = workbook.Sheets['BOM'];
        bomJson = XLSX.utils.sheet_to_json(bomSheet);
        setBomData(bomJson);
      }

      if (workbook.SheetNames.includes('CAMPAIGN')) {
        const campaignSheet = workbook.Sheets['CAMPAIGN'];
        campaignJson = XLSX.utils.sheet_to_json(campaignSheet, { raw: false });
        campaignJson = campaignJson.map(row => {
          if (row.DATE) {
            const date = new Date(row.DATE);
            const monthName = date.toLocaleString('en-US', { month: 'long' }).toUpperCase();
            return { ...row, MONTH: monthName };
          }
          return row;
        });
        setCampaignData(campaignJson);
      }

      if (workbook.SheetNames.includes('CAMPAIGN_BOM')) {
        const campaignBomSheet = workbook.Sheets['CAMPAIGN_BOM'];
        campaignBomJson = XLSX.utils.sheet_to_json(campaignBomSheet);
        setCampaignBomData(campaignBomJson);
      }

      // Load Field Result Tracker sheets
      if (workbook.SheetNames.includes('FIELD_DAILY')) {
        const fieldDailySheet = workbook.Sheets['FIELD_DAILY'];
        fieldDailyJson = XLSX.utils.sheet_to_json(fieldDailySheet, { raw: false });
        fieldDailyJson = fieldDailyJson.map(row => {
          if (row.DATE) {
            const date = new Date(row.DATE);
            const monthName = date.toLocaleString('en-US', { month: 'long' }).toUpperCase();
            return { ...row, MONTH: monthName };
          }
          return row;
        });
        setFieldDailyData(fieldDailyJson);
      }

      if (workbook.SheetNames.includes('FIELD_BOM')) {
        const fieldBomSheet = workbook.Sheets['FIELD_BOM'];
        fieldBomJson = XLSX.utils.sheet_to_json(fieldBomSheet);
        setFieldBomData(fieldBomJson);
      }

      if (workbook.SheetNames.includes('FIELD_CAMPAIGN')) {
        const fieldCampaignSheet = workbook.Sheets['FIELD_CAMPAIGN'];
        fieldCampaignJson = XLSX.utils.sheet_to_json(fieldCampaignSheet, { raw: false });
        fieldCampaignJson = fieldCampaignJson.map(row => {
          if (row.DATE) {
            const date = new Date(row.DATE);
            const monthName = date.toLocaleString('en-US', { month: 'long' }).toUpperCase();
            return { ...row, MONTH: monthName };
          }
          return row;
        });
        setFieldCampaignData(fieldCampaignJson);
      }

      if (workbook.SheetNames.includes('PER_AREA')) {
        const perAreaSheet = workbook.Sheets['PER_AREA'];
        perAreaJson = XLSX.utils.sheet_to_json(perAreaSheet, { raw: false });
        perAreaJson = perAreaJson.map(row => {
          if (row.DATE) {
            const date = new Date(row.DATE);
            const monthName = date.toLocaleString('en-US', { month: 'long' }).toUpperCase();
            return { ...row, MONTH: monthName };
          }
          return row;
        });
        setPerAreaData(perAreaJson);
      }

      const months = [...new Set([...dailyJson.map(row => row.MONTH), ...campaignJson.map(row => row.MONTH), ...fieldDailyJson.map(row => row.MONTH)].filter(Boolean))];
      const productTypes = [...new Set([...dailyJson.map(row => row['PRODUCT TYPE']), ...campaignJson.map(row => row['PRODUCT TYPE']), ...fieldDailyJson.map(row => row['PRODUCT TYPE'])].filter(Boolean))].sort();
      const clients = [...new Set([...campaignJson.map(row => row.CAMPAIGN), ...fieldCampaignJson.map(row => row.CAMPAIGN)].filter(Boolean))].sort();
      const areas = [...new Set(perAreaJson.map(row => row.AREA).filter(Boolean))].sort();
      
      setAvailableMonths(months);
      setAvailableProductTypes(productTypes);
      setAvailableClients(clients);
      setAvailableAreas(areas);
      
      if (months.length > 0) {
        setSelectedMonth(months[0]);
        setSelectedMonths([months[0]]);
        setSelectedComparisonProductType(productTypes[0] || '');
      }
      if (productTypes.length > 0) setSelectedProductType(productTypes[0]);
      if (clients.length > 0) setSelectedClient(clients[0]);
    } catch (error) {
      setError(error.message || 'Error reading file');
    }
    setLoading(false);
  };

  const resetFilters = () => {
    setStartDate('');
    setEndDate('');
    setFieldStartDate('');
    setFieldEndDate('');
  };

  const filteredOverallData = useMemo(() => {
    return dailyData.filter(row => {
      if (row.MONTH !== selectedMonth || row['PRODUCT TYPE'] !== selectedProductType) return false;
      
      if (row.DATE && (startDate || endDate)) {
        const rowDate = normalizeDateForComparison(row.DATE);
        if (startDate && endDate) {
          const start = normalizeDateForComparison(startDate);
          const end = new Date(endDate);
          end.setHours(23, 59, 59, 999);
          return rowDate >= start && rowDate <= end;
        } else if (startDate) {
          return rowDate >= normalizeDateForComparison(startDate);
        } else if (endDate) {
          const end = new Date(endDate);
          end.setHours(23, 59, 59, 999);
          return rowDate <= end;
        }
      }
      return true;
    });
  }, [dailyData, selectedMonth, selectedProductType, startDate, endDate]);

  const filteredClientData = useMemo(() => {
    return campaignData.filter(row => {
      if (row.MONTH !== selectedMonth || row.CAMPAIGN !== selectedClient) return false;
      
      if (row.DATE && (startDate || endDate)) {
        const rowDate = normalizeDateForComparison(row.DATE);
        if (startDate && endDate) {
          const start = normalizeDateForComparison(startDate);
          const end = new Date(endDate);
          end.setHours(23, 59, 59, 999);
          return rowDate >= start && rowDate <= end;
        } else if (startDate) {
          return rowDate >= normalizeDateForComparison(startDate);
        } else if (endDate) {
          const end = new Date(endDate);
          end.setHours(23, 59, 59, 999);
          return rowDate <= end;
        }
      }
      return true;
    });
  }, [campaignData, selectedMonth, selectedClient, startDate, endDate]);

  const filteredFieldData = useMemo(() => {
    return fieldDailyData.filter(row => {
      if (row.MONTH !== selectedMonth || row['PRODUCT TYPE'] !== selectedProductType) return false;
      
      if (row.DATE && (fieldStartDate || fieldEndDate)) {
        const rowDate = normalizeDateForComparison(row.DATE);
        if (fieldStartDate && fieldEndDate) {
          const start = normalizeDateForComparison(fieldStartDate);
          const end = new Date(fieldEndDate);
          end.setHours(23, 59, 59, 999);
          return rowDate >= start && rowDate <= end;
        } else if (fieldStartDate) {
          return rowDate >= normalizeDateForComparison(fieldStartDate);
        } else if (fieldEndDate) {
          const end = new Date(fieldEndDate);
          end.setHours(23, 59, 59, 999);
          return rowDate <= end;
        }
      }
      return true;
    });
  }, [fieldDailyData, selectedMonth, selectedProductType, fieldStartDate, fieldEndDate]);

  const filteredFieldCampaignData = useMemo(() => {
    return fieldCampaignData.filter(row => {
      if (row.MONTH !== selectedMonth || row['PRODUCT TYPE'] !== selectedProductType) return false;
      
      if (row.DATE && (fieldStartDate || fieldEndDate)) {
        const rowDate = normalizeDateForComparison(row.DATE);
        if (fieldStartDate && fieldEndDate) {
          const start = normalizeDateForComparison(fieldStartDate);
          const end = new Date(fieldEndDate);
          end.setHours(23, 59, 59, 999);
          return rowDate >= start && rowDate <= end;
        } else if (fieldStartDate) {
          return rowDate >= normalizeDateForComparison(fieldStartDate);
        } else if (fieldEndDate) {
          const end = new Date(fieldEndDate);
          end.setHours(23, 59, 59, 999);
          return rowDate <= end;
        }
      }
      return true;
    });
  }, [fieldCampaignData, selectedMonth, selectedProductType, fieldStartDate, fieldEndDate]);

  const filteredPerAreaData = useMemo(() => {
    return perAreaData.filter(row => {
      // Check product type
      if (row['PRODUCT TYPE'] !== selectedProductType) return false;
      
      // Check if the date is in the selected month
      if (row.DATE) {
        const date = new Date(row.DATE);
        const rowMonth = date.toLocaleString('en-US', { month: 'long' }).toUpperCase();
        if (rowMonth !== selectedMonth) return false;
        
        // Apply date filters
        if (fieldStartDate || fieldEndDate) {
          const rowDate = normalizeDateForComparison(row.DATE);
          if (fieldStartDate && fieldEndDate) {
            const start = normalizeDateForComparison(fieldStartDate);
            const end = new Date(fieldEndDate);
            end.setHours(23, 59, 59, 999);
            return rowDate >= start && rowDate <= end;
          } else if (fieldStartDate) {
            return rowDate >= normalizeDateForComparison(fieldStartDate);
          } else if (fieldEndDate) {
            const end = new Date(fieldEndDate);
            end.setHours(23, 59, 59, 999);
            return rowDate <= end;
          }
        }
      }
      return true;
    });
  }, [perAreaData, selectedMonth, selectedProductType, fieldStartDate, fieldEndDate]);

  // Filtered data for Field Result Per Campaign (by client)
  const filteredFieldCampaignByClient = useMemo(() => {
    return fieldCampaignData.filter(row => {
      if (row.MONTH !== selectedMonth || row.CAMPAIGN !== selectedClient) return false;
      
      if (row.DATE && (fieldStartDate || fieldEndDate)) {
        const rowDate = normalizeDateForComparison(row.DATE);
        if (fieldStartDate && fieldEndDate) {
          const start = normalizeDateForComparison(fieldStartDate);
          const end = new Date(fieldEndDate);
          end.setHours(23, 59, 59, 999);
          return rowDate >= start && rowDate <= end;
        } else if (fieldStartDate) {
          return rowDate >= normalizeDateForComparison(fieldStartDate);
        } else if (fieldEndDate) {
          const end = new Date(fieldEndDate);
          end.setHours(23, 59, 59, 999);
          return rowDate <= end;
        }
      }
      return true;
    });
  }, [fieldCampaignData, selectedMonth, selectedClient, fieldStartDate, fieldEndDate]);

  const filteredPerAreaByClient = useMemo(() => {
    return perAreaData.filter(row => {
      // Check campaign
      if (row.CAMPAIGN !== selectedClient) return false;
      
      // Check if the date is in the selected month
      if (row.DATE) {
        const date = new Date(row.DATE);
        const rowMonth = date.toLocaleString('en-US', { month: 'long' }).toUpperCase();
        if (rowMonth !== selectedMonth) return false;
        
        // Apply date filters
        if (fieldStartDate || fieldEndDate) {
          const rowDate = normalizeDateForComparison(row.DATE);
          if (fieldStartDate && fieldEndDate) {
            const start = normalizeDateForComparison(fieldStartDate);
            const end = new Date(fieldEndDate);
            end.setHours(23, 59, 59, 999);
            return rowDate >= start && rowDate <= end;
          } else if (fieldStartDate) {
            return rowDate >= normalizeDateForComparison(fieldStartDate);
          } else if (fieldEndDate) {
            const end = new Date(fieldEndDate);
            end.setHours(23, 59, 59, 999);
            return rowDate <= end;
          }
        }
      }
      return true;
    });
  }, [perAreaData, selectedMonth, selectedClient, fieldStartDate, fieldEndDate]);

  const fieldMetrics = useMemo(() => {
    if (filteredFieldData.length === 0) return null;
    
    // Get Field BOM from FIELD_BOM sheet
    const bomRecord = fieldBomData.find(r => r.MONTH === selectedMonth && r['PRODUCT TYPE'] === selectedProductType);
    const fieldBom = parseNumber(bomRecord?.BOM_COUNT);
    
    // Get total new endorsements from DAILY sheet for the selected month (MTD based on filters)
    const monthlyDailyData = dailyData.filter(r => {
      if (r.MONTH !== selectedMonth || r['PRODUCT TYPE'] !== selectedProductType) return false;
      
      // Apply same date filters as field data for MTD calculation
      if (r.DATE && (fieldStartDate || fieldEndDate)) {
        const rowDate = normalizeDateForComparison(r.DATE);
        if (fieldStartDate && fieldEndDate) {
          const start = normalizeDateForComparison(fieldStartDate);
          const end = new Date(fieldEndDate);
          end.setHours(23, 59, 59, 999);
          return rowDate >= start && rowDate <= end;
        } else if (fieldStartDate) {
          return rowDate >= normalizeDateForComparison(fieldStartDate);
        } else if (fieldEndDate) {
          const end = new Date(fieldEndDate);
          end.setHours(23, 59, 59, 999);
          return rowDate <= end;
        }
      }
      return true;
    });
    
    const totalNewEndo = monthlyDailyData.reduce((sum, row) => {
      return sum + parseNumber(row[viewMode === 'ob' ? 'ENDORSEMENTS OB' : 'ENDORSEMENTS']);
    }, 0);
    
    // GROSS = Field BOM + Total New Endorsements
    const gross = fieldBom + totalNewEndo;
    
    // VISITED = Sum of all daily TNA values (total accounts visited MTD)
    // TNA represents daily visit count, so we sum all days
    const totalVisited = filteredFieldData.reduce((sum, row) => {
      return sum + parseNumber(row.TNA);
    }, 0);
    
    // PENDING = GROSS - VISITED
    const totalPending = gross - totalVisited;
    
    return { gross, fieldBom, totalVisited, totalPending };
  }, [filteredFieldData, fieldBomData, dailyData, selectedMonth, selectedProductType, viewMode, fieldStartDate, fieldEndDate]);

  // Field metrics for campaign-specific view
  const fieldCampaignMetrics = useMemo(() => {
    if (filteredFieldCampaignByClient.length === 0) return null;
    
    // Get Field BOM from FIELD_BOM sheet for this campaign
    const bomRecord = fieldBomData.find(r => r.MONTH === selectedMonth && r.CAMPAIGN === selectedClient);
    const fieldBom = parseNumber(bomRecord?.BOM_COUNT);
    
    // Get total new endorsements from CAMPAIGN sheet for the selected month (MTD based on filters)
    const monthlyCampaignData = campaignData.filter(r => {
      if (r.MONTH !== selectedMonth || r.CAMPAIGN !== selectedClient) return false;
      
      // Apply same date filters as field data for MTD calculation
      if (r.DATE && (fieldStartDate || fieldEndDate)) {
        const rowDate = normalizeDateForComparison(r.DATE);
        if (fieldStartDate && fieldEndDate) {
          const start = normalizeDateForComparison(fieldStartDate);
          const end = new Date(fieldEndDate);
          end.setHours(23, 59, 59, 999);
          return rowDate >= start && rowDate <= end;
        } else if (fieldStartDate) {
          return rowDate >= normalizeDateForComparison(fieldStartDate);
        } else if (fieldEndDate) {
          const end = new Date(fieldEndDate);
          end.setHours(23, 59, 59, 999);
          return rowDate <= end;
        }
      }
      return true;
    });
    
    const totalNewEndo = monthlyCampaignData.reduce((sum, row) => {
      return sum + parseNumber(row[viewMode === 'ob' ? 'NEW ENDO OB' : 'NEW ENDO']);
    }, 0);
    
    // GROSS = Field BOM + Total New Endorsements
    const gross = fieldBom + totalNewEndo;
    
    // VISITED = Sum of all daily TNA values from FIELD_CAMPAIGN
    const totalVisited = filteredFieldCampaignByClient.reduce((sum, row) => {
      return sum + parseNumber(row.TNA);
    }, 0);
    
    // PENDING = GROSS - VISITED
    const totalPending = gross - totalVisited;
    
    return { gross, fieldBom, totalVisited, totalPending };
  }, [filteredFieldCampaignByClient, fieldBomData, campaignData, selectedMonth, selectedClient, viewMode, fieldStartDate, fieldEndDate]);

  const fieldDailyChartData = useMemo(() => {
    const sorted = [...filteredFieldData].sort((a, b) => new Date(a.DATE) - new Date(b.DATE));
    
    return sorted.map(row => {
      // TNA is daily visit count - show it as is (not cumulative)
      const dailyVisits = parseNumber(row.TNA);
      
      return {
        date: row.DATE || '',
        visited: dailyVisits
      };
    });
  }, [filteredFieldData]);

  const fieldCampaignChartData = useMemo(() => {
    const sorted = [...filteredFieldCampaignByClient].sort((a, b) => new Date(a.DATE) - new Date(b.DATE));
    
    return sorted.map(row => {
      const dailyVisits = parseNumber(row.TNA);
      
      return {
        date: row.DATE || '',
        visited: dailyVisits
      };
    });
  }, [filteredFieldCampaignByClient]);

  const fieldPerClientChartData = useMemo(() => {
    const clientMap = new Map();
    
    filteredFieldCampaignData.forEach(row => {
      const client = row.CAMPAIGN;
      if (!client) return;
      
      if (!clientMap.has(client)) {
        clientMap.set(client, []);
      }
      
      clientMap.get(client).push({
        date: row.DATE,
        visited: parseNumber(row.TNA)
      });
    });
    
    // Get all unique dates
    const allDates = [...new Set(filteredFieldCampaignData.map(r => r.DATE))].sort((a, b) => new Date(a) - new Date(b));
    
    // Build chart data with each client as a series
    return allDates.map(date => {
      const dataPoint = { date };
      clientMap.forEach((records, client) => {
        const record = records.find(r => r.date === date);
        dataPoint[client] = record ? record.visited : 0;
      });
      return dataPoint;
    });
  }, [filteredFieldCampaignData]);

  const fieldPerAreaChartData = useMemo(() => {
    if (filteredPerAreaData.length === 0) return [];
    
    // Group data by Area and Date, summing TNA across all campaigns in that area
    const areaDateMap = new Map();
    
    filteredPerAreaData.forEach(row => {
      const area = row.AREA;
      const date = row.DATE;
      if (!area || !date) return;
      
      const key = `${area}|${date}`;
      if (!areaDateMap.has(key)) {
        areaDateMap.set(key, {
          area,
          date,
          visited: 0
        });
      }
      
      // Sum TNA for all campaigns in this area on this date
      const current = areaDateMap.get(key);
      current.visited += parseNumber(row.TNA);
    });
    
    // Get unique dates and areas
    const allDates = [...new Set(filteredPerAreaData.map(r => r.DATE).filter(Boolean))].sort((a, b) => new Date(a) - new Date(b));
    const allAreas = [...new Set(filteredPerAreaData.map(r => r.AREA).filter(Boolean))];
    
    // Build chart data
    return allDates.map(date => {
      const dataPoint = { date };
      allAreas.forEach(area => {
        const key = `${area}|${date}`;
        const data = areaDateMap.get(key);
        dataPoint[area] = data ? data.visited : 0;
      });
      return dataPoint;
    });
  }, [filteredPerAreaData]);

  const fieldPerAreaByClientChartData = useMemo(() => {
    if (filteredPerAreaByClient.length === 0) return [];
    
    // Group data by Area and Date
    const areaDateMap = new Map();
    
    filteredPerAreaByClient.forEach(row => {
      const area = row.AREA;
      const date = row.DATE;
      if (!area || !date) return;
      
      const key = `${area}|${date}`;
      if (!areaDateMap.has(key)) {
        areaDateMap.set(key, {
          area,
          date,
          visited: 0
        });
      }
      
      const current = areaDateMap.get(key);
      current.visited += parseNumber(row.TNA);
    });
    
    // Get unique dates and areas for this client
    const allDates = [...new Set(filteredPerAreaByClient.map(r => r.DATE).filter(Boolean))].sort((a, b) => new Date(a) - new Date(b));
    const allAreas = [...new Set(filteredPerAreaByClient.map(r => r.AREA).filter(Boolean))];
    
    // Build chart data
    return allDates.map(date => {
      const dataPoint = { date };
      allAreas.forEach(area => {
        const key = `${area}|${date}`;
        const data = areaDateMap.get(key);
        dataPoint[area] = data ? data.visited : 0;
      });
      return dataPoint;
    });
  }, [filteredPerAreaByClient]);

  const clientNames = useMemo(() => {
    const clients = new Set();
    filteredFieldCampaignData.forEach(row => {
      const client = row.CAMPAIGN;
      if (client) clients.add(client);
    });
    return Array.from(clients);
  }, [filteredFieldCampaignData]);

  const areaNames = useMemo(() => {
    return availableAreas;
  }, [availableAreas]);

  const areaNamesForClient = useMemo(() => {
    const areas = new Set();
    filteredPerAreaByClient.forEach(row => {
      if (row.AREA) areas.add(row.AREA);
    });
    return Array.from(areas);
  }, [filteredPerAreaByClient]);

  const overallMetrics = useMemo(() => {
    if (filteredOverallData.length === 0) return null;
    const bomRecord = bomData.find(r => r.MONTH === selectedMonth && r['PRODUCT TYPE'] === selectedProductType);
    const bom = parseNumber(bomRecord?.[viewMode === 'ob' ? 'OB' : 'TNA']);
    const sorted = [...filteredOverallData].sort((a, b) => new Date(a.DATE) - new Date(b.DATE));
    const active = parseNumber(sorted[sorted.length - 1]?.[viewMode === 'ob' ? 'Total Portfolio OB' : 'Total Portfolio']);
    const totalEndorsements = filteredOverallData.reduce((sum, row) => sum + parseNumber(row[viewMode === 'ob' ? 'ENDORSEMENTS OB' : 'ENDORSEMENTS']), 0);
    const totalPullouts = filteredOverallData.reduce((sum, row) => sum + parseNumber(row[viewMode === 'ob' ? 'PULLOUT OB' : 'PULLOUT']), 0);
    const portfolioGrowth = bom !== 0 ? ((active - bom) / bom) * 100 : 0;
    const netFlowObj = calculateNetFlow(totalEndorsements, totalPullouts);
    return { bom, active, portfolioGrowth, netFlowObj, totalEndorsements, totalPullouts };
  }, [filteredOverallData, bomData, selectedMonth, selectedProductType, viewMode]);

  const clientMetrics = useMemo(() => {
    if (filteredClientData.length === 0) return null;
    const bomRecord = campaignBomData.find(r => r.MONTH === selectedMonth && r.CAMPAIGN === selectedClient);
    const bom = parseNumber(bomRecord?.[viewMode === 'ob' ? 'OB' : 'TNA']);
    const sorted = [...filteredClientData].sort((a, b) => new Date(a.DATE) - new Date(b.DATE));
    const active = parseNumber(sorted[sorted.length - 1]?.[viewMode === 'ob' ? 'Total Portfolio OB' : 'Total Portfolio']);
    const totalEndorsements = filteredClientData.reduce((sum, row) => sum + parseNumber(row[viewMode === 'ob' ? 'NEW ENDO OB' : 'NEW ENDO']), 0);
    const totalPullouts = filteredClientData.reduce((sum, row) => sum + parseNumber(row[viewMode === 'ob' ? 'PULLOUT OB' : 'PULLOUT']), 0);
    const portfolioGrowth = bom !== 0 ? ((active - bom) / bom) * 100 : 0;
    const netFlowObj = calculateNetFlow(totalEndorsements, totalPullouts);
    return { bom, active, portfolioGrowth, netFlowObj, totalEndorsements, totalPullouts };
  }, [filteredClientData, campaignBomData, selectedMonth, selectedClient, viewMode]);

  const getChartData = useMemo(() => {
    const data = activeTab === 'overall' ? filteredOverallData : filteredClientData;
    const metrics = activeTab === 'overall' ? overallMetrics : clientMetrics;
    const sorted = [...data].sort((a, b) => new Date(a.DATE) - new Date(b.DATE));
    
    return sorted.map((row, index) => {
      const endorsements = parseNumber(row[activeTab === 'overall' 
        ? (viewMode === 'ob' ? 'ENDORSEMENTS OB' : 'ENDORSEMENTS')
        : (viewMode === 'ob' ? 'NEW ENDO OB' : 'NEW ENDO')
      ]);
      const pullouts = parseNumber(row[activeTab === 'overall'
        ? (viewMode === 'ob' ? 'PULLOUT OB' : 'PULLOUT')
        : (viewMode === 'ob' ? 'PULLOUT OB' : 'PULLOUT')
      ]);
      const portfolio = parseNumber(row[viewMode === 'ob' ? 'Total Portfolio OB' : 'Total Portfolio']);
      
      const netFlowRatio = pullouts === 0 ? 0 : endorsements / pullouts;
      const portfolioGrowth = metrics && metrics.bom !== 0 ? ((portfolio - metrics.bom) / metrics.bom) * 100 : 0;
      
      return {
        date: row.DATE || `Day ${index + 1}`,
        endorsements,
        pullouts,
        netFlowRatio,
        portfolioGrowth,
        portfolio
      };
    });
  }, [activeTab, filteredOverallData, filteredClientData, overallMetrics, clientMetrics, viewMode]);

  const getMTDData = useMemo(() => {
    return getChartData.map((row, index) => {
      const cumulativeEndorsements = getChartData.slice(0, index + 1).reduce((sum, r) => sum + r.endorsements, 0);
      const cumulativePullouts = getChartData.slice(0, index + 1).reduce((sum, r) => sum + r.pullouts, 0);
      const mtdNetFlowRatio = cumulativePullouts === 0 ? 0 : cumulativeEndorsements / cumulativePullouts;
      const metrics = activeTab === 'overall' ? overallMetrics : clientMetrics;
      const mtdPortfolioGrowth = metrics && metrics.bom !== 0 ? ((row.portfolio - metrics.bom) / metrics.bom) * 100 : 0;
      
      return {
        date: row.date,
        cumulativeEndorsements,
        cumulativePullouts,
        mtdNetFlowRatio,
        mtdPortfolioGrowth
      };
    });
  }, [getChartData, activeTab, overallMetrics, clientMetrics]);

  const monthlyComparisonData = useMemo(() => {
    if (activeTab !== 'monthly') return { monthlyMetrics: [], clientComparison: [] };

    const maxDay = Math.min(...selectedMonths.map(month => {
      const monthData = dailyData.filter(r => r.MONTH === month && r['PRODUCT TYPE'] === selectedComparisonProductType);
      if (monthData.length === 0) return 31;
      return Math.max(...monthData.map(r => r.DATE ? new Date(r.DATE).getDate() : 0));
    }));

    const monthlyMetrics = selectedMonths.map(month => {
      const monthData = dailyData.filter(r => {
        if (r.MONTH !== month || r['PRODUCT TYPE'] !== selectedComparisonProductType) return false;
        if (r.DATE) {
          const day = new Date(r.DATE).getDate();
          if (day > maxDay) return false;
        }
        return true;
      });

      const bomRecord = bomData.find(r => r.MONTH === month && r['PRODUCT TYPE'] === selectedComparisonProductType);
      const bom = parseNumber(bomRecord?.[viewMode === 'ob' ? 'OB' : 'TNA']);
      
      const sorted = [...monthData].sort((a, b) => new Date(a.DATE) - new Date(b.DATE));
      const active = parseNumber(sorted[sorted.length - 1]?.[viewMode === 'ob' ? 'Total Portfolio OB' : 'Total Portfolio']);
      
      const totalEndorsements = monthData.reduce((sum, r) => sum + parseNumber(r[viewMode === 'ob' ? 'ENDORSEMENTS OB' : 'ENDORSEMENTS']), 0);
      const totalPullouts = monthData.reduce((sum, r) => sum + parseNumber(r[viewMode === 'ob' ? 'PULLOUT OB' : 'PULLOUT']), 0);
      
      const portfolioGrowth = bom !== 0 ? ((active - bom) / bom) * 100 : 0;
      const netFlowObj = calculateNetFlow(totalEndorsements, totalPullouts);

      return { month, bom, active, portfolioGrowth, netFlowObj, totalEndorsements, totalPullouts };
    });

    const currentMonth = selectedMonths[0];
    const clientMap = new Map();

    selectedMonths.forEach(month => {
      campaignData.filter(r => {
        if (r.MONTH !== month) return false;
        if (r.DATE) {
          const day = new Date(r.DATE).getDate();
          if (day > maxDay) return false;
        }
        return true;
      }).forEach(r => {
        const client = r.CAMPAIGN;
        if (!client) return;
        if (!clientMap.has(client)) clientMap.set(client, { totalEndorsements: 0, totalPullouts: 0 });
        const data = clientMap.get(client);
        data.totalEndorsements += parseNumber(r[viewMode === 'ob' ? 'NEW ENDO OB' : 'NEW ENDO']);
        data.totalPullouts += parseNumber(r[viewMode === 'ob' ? 'PULLOUT OB' : 'PULLOUT']);
      });
    });

    const clientComparison = Array.from(clientMap.entries()).map(([name, data]) => {
      const bomRecord = campaignBomData.find(r => r.CAMPAIGN === name && r.MONTH === currentMonth);
      const bom = parseNumber(bomRecord?.[viewMode === 'ob' ? 'OB' : 'TNA']);
      
      const lastData = campaignData.filter(r => {
        if (r.CAMPAIGN !== name || r.MONTH !== currentMonth) return false;
        if (r.DATE) {
          const day = new Date(r.DATE).getDate();
          if (day > maxDay) return false;
        }
        return true;
      }).sort((a, b) => new Date(b.DATE) - new Date(a.DATE))[0];
      
      const active = parseNumber(lastData?.[viewMode === 'ob' ? 'Total Portfolio OB' : 'Total Portfolio']);
      const portfolioGrowth = bom !== 0 ? ((active - bom) / bom) * 100 : 0;
      const netFlowObj = calculateNetFlow(data.totalEndorsements, data.totalPullouts);

      return { name, bom, active, portfolioGrowth, netFlowObj, ...data };
    }).sort((a, b) => b.portfolioGrowth - a.portfolioGrowth);

    const filteredClients = clientRankingView === 'top5' 
      ? clientComparison.slice(0, 5)
      : clientRankingView === 'bottom5'
      ? clientComparison.slice(-5).reverse()
      : clientComparison;

    return { monthlyMetrics, clientComparison: filteredClients, maxDay };
  }, [activeTab, selectedMonths, dailyData, campaignData, bomData, campaignBomData, selectedComparisonProductType, viewMode, clientRankingView]);

  const renderWelcome = () => (
    <div className="flex flex-col items-center justify-center py-20">
      <h2 className="text-3xl font-bold text-gray-800 mb-4">Welcome to MC03 Endorsement Flow Monitor</h2>
      <p className="text-gray-600 mb-8">Upload your Excel file to start analyzing endorsement data</p>
      <div className="bg-blue-50 border border-blue-200 rounded-lg p-6 max-w-md">
        <h3 className="font-semibold text-gray-800 mb-3">📋 Required Sheets:</h3>
        <ul className="space-y-2 text-sm text-gray-700">
          <li>• DAILY - Daily endorsement data</li>
          <li>• BOM - Beginning of month values</li>
          <li>• CAMPAIGN - Client-specific data</li>
          <li>• CAMPAIGN_BOM - Campaign BOM values</li>
          <li>• FIELD_DAILY - Daily field visit data (optional)</li>
          <li>• FIELD_BOM - Field BOM values (optional)</li>
          <li>• FIELD_CAMPAIGN - Campaign field data (optional)</li>
          <li>• PER_AREA - Area-specific data (optional)</li>
        </ul>
      </div>
    </div>
  );

  const getRandomColor = (index) => {
    const colors = [
      '#3b82f6', '#ef4444', '#10b981', '#f59e0b', '#8b5cf6',
      '#ec4899', '#14b8a6', '#f97316', '#6366f1', '#84cc16'
    ];
    return colors[index % colors.length];
  };

  if (dailyData.length === 0) {
    return (
      <div className="min-h-screen bg-white p-6">
        <div className="max-w-7xl mx-auto">
          <div className="flex justify-between items-center mb-6 border-b-2 border-gray-300 pb-4">
            <h1 className="text-2xl font-bold">MC03 Endorsement Flow Monitoring</h1>
            <label className="px-4 py-2 bg-indigo-600 text-white rounded cursor-pointer hover:bg-indigo-700">
              Upload Excel File
              <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} className="hidden" />
            </label>
          </div>
          {loading && <div className="text-center py-8 text-indigo-600">Loading...</div>}
          {error && (
            <div className="bg-red-50 border border-red-200 rounded p-4 mb-6 flex gap-3">
              <AlertCircle className="h-5 w-5 text-red-600 flex-shrink-0" />
              <div>
                <h3 className="font-semibold text-red-800">Error Loading File</h3>
                <p className="text-sm text-red-700 whitespace-pre-line">{error}</p>
              </div>
            </div>
          )}
          {!loading && !error && renderWelcome()}
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-white p-6">
      <div className="max-w-7xl mx-auto">
        <div className="flex justify-between items-center mb-4 border-b-2 border-gray-300 pb-4">
          <h1 className="text-2xl font-bold">MC03 Endorsement Flow Monitoring</h1>
          <div className="flex gap-2">
            <button onClick={resetFilters} className="px-3 py-2 border border-gray-300 rounded hover:bg-gray-50 flex items-center gap-2">
              <RefreshCw className="h-4 w-4" />
              Reset
            </button>
            <label className="px-4 py-2 bg-indigo-600 text-white rounded cursor-pointer hover:bg-indigo-700">
              Upload New
              <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} className="hidden" />
            </label>
          </div>
        </div>

        <div className="flex gap-4 mb-6 border-b border-gray-200">
          <button onClick={() => setActiveTab('overall')} className={`px-4 py-2 font-medium transition-colors ${activeTab === 'overall' ? 'border-b-2 border-indigo-600 text-indigo-600' : 'text-gray-500 hover:text-gray-700'}`}>Overall</button>
          <button onClick={() => setActiveTab('client')} className={`px-4 py-2 font-medium transition-colors ${activeTab === 'client' ? 'border-b-2 border-indigo-600 text-indigo-600' : 'text-gray-500 hover:text-gray-700'}`}>Client</button>
          <button onClick={() => setActiveTab('monthly')} className={`px-4 py-2 font-medium transition-colors ${activeTab === 'monthly' ? 'border-b-2 border-indigo-600 text-indigo-600' : 'text-gray-500 hover:text-gray-700'}`}>Monthly Comparison</button>
          <button onClick={() => setActiveTab('field')} className={`px-4 py-2 font-medium transition-colors ${activeTab === 'field' ? 'border-b-2 border-indigo-600 text-indigo-600' : 'text-gray-500 hover:text-gray-700'}`}>Field Result Tracker</button>
          <button onClick={() => setActiveTab('fieldCampaign')} className={`px-4 py-2 font-medium transition-colors ${activeTab === 'fieldCampaign' ? 'border-b-2 border-indigo-600 text-indigo-600' : 'text-gray-500 hover:text-gray-700'}`}>Field Result Per Campaign</button>
        </div>

        {activeTab === 'field' ? (
          <>
            <div className="bg-white rounded-lg p-6 shadow-sm border mb-6">
              <div className="grid grid-cols-4 gap-4">
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-2">Month</label>
                  <select value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)} className="w-full p-2 border border-gray-300 rounded text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500">
                    {availableMonths.map(m => <option key={m} value={m}>{m}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-2">Product Type</label>
                  <select value={selectedProductType} onChange={(e) => setSelectedProductType(e.target.value)} className="w-full p-2 border border-gray-300 rounded text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500">
                    {availableProductTypes.map(t => <option key={t} value={t}>{t}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-2">Start Date</label>
                  <input type="date" value={fieldStartDate} onChange={(e) => setFieldStartDate(e.target.value)} className="w-full p-2 border border-gray-300 rounded text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500" />
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-2">End Date</label>
                  <input type="date" value={fieldEndDate} onChange={(e) => setFieldEndDate(e.target.value)} className="w-full p-2 border border-gray-300 rounded text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500" />
                </div>
              </div>
            </div>

            {fieldMetrics && (
              <div className="grid grid-cols-4 gap-6 mb-6">
                <div className="bg-white border rounded-lg p-6 shadow-sm">
                  <div className="text-sm font-semibold text-gray-600 mb-2">Total Portfolio</div>
                  <div className="text-3xl font-bold text-gray-900">{formatNumber(fieldMetrics.gross)}</div>
                  <div className="text-xs text-gray-500 mt-2">
                    BOM: {formatNumber(fieldMetrics.fieldBom)} + New Endo: {formatNumber(fieldMetrics.gross - fieldMetrics.fieldBom)}
                  </div>
                </div>
                <div className="bg-white border rounded-lg p-6 shadow-sm">
                  <div className="text-sm font-semibold text-gray-600 mb-2">Field BOM</div>
                  <div className="text-3xl font-bold text-gray-900">{formatNumber(fieldMetrics.fieldBom)}</div>
                  <div className="text-xs text-gray-500 mt-2">
                    Beginning of month accounts
                  </div>
                </div>
                <div className="bg-white border rounded-lg p-6 shadow-sm">
                  <div className="text-sm font-semibold text-gray-600 mb-2">VISITED</div>
                  <div className="text-3xl font-bold text-green-600">{formatNumber(fieldMetrics.totalVisited)}</div>
                  <div className="text-xs text-gray-500 mt-2">
                    Total accounts visited (MTD)
                  </div>
                </div>
                <div className="bg-white border rounded-lg p-6 shadow-sm">
                  <div className="text-sm font-semibold text-gray-600 mb-2">PENDING</div>
                  <div className="text-3xl font-bold text-orange-600">{formatNumber(fieldMetrics.totalPending)}</div>
                  <div className="text-xs text-gray-500 mt-2">
                    Gross: {formatNumber(fieldMetrics.gross)} - Visited: {formatNumber(fieldMetrics.totalVisited)}
                  </div>
                </div>
              </div>
            )}

            <div className="bg-white rounded-lg p-6 shadow-sm border mb-6">
              <h3 className="text-sm font-semibold text-gray-700 mb-4 uppercase">{selectedProductType} - DAILY VISITATION SUMMARY</h3>
              <div className="border rounded-lg p-4 bg-gray-50" style={{ minHeight: '350px' }}>
                <ResponsiveContainer width="100%" height={300}>
                  <LineChart data={fieldDailyChartData}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="date" tick={{ fontSize: 10 }} angle={-45} textAnchor="end" height={60} />
                    <YAxis tick={{ fontSize: 10 }} />
                    <RechartsTooltip />
                    <Legend />
                    <Line type="monotone" dataKey="visited" stroke="#10b981" strokeWidth={2} name="Visited">
                      <LabelList dataKey="visited" position="top" style={{ fontSize: '11px', fontWeight: 'bold', fill: '#10b981' }} />
                    </Line>
                  </LineChart>
                </ResponsiveContainer>
              </div>
            </div>

            <div className="bg-white rounded-lg p-6 shadow-sm border mb-6">
              <h3 className="text-sm font-semibold text-gray-700 mb-4 uppercase">PER CLIENT</h3>
              <div className="border rounded-lg p-4 bg-gray-50" style={{ minHeight: '350px' }}>
                <ResponsiveContainer width="100%" height={300}>
                  <LineChart data={fieldPerClientChartData}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="date" tick={{ fontSize: 10 }} angle={-45} textAnchor="end" height={60} />
                    <YAxis tick={{ fontSize: 10 }} />
                    <RechartsTooltip />
                    <Legend />
                    {clientNames.map((client, index) => (
                      <Line 
                        key={client} 
                        type="monotone" 
                        dataKey={client} 
                        stroke={getRandomColor(index)} 
                        strokeWidth={2} 
                      />
                    ))}
                  </LineChart>
                </ResponsiveContainer>
              </div>
            </div>

            <div className="bg-white rounded-lg p-6 shadow-sm border">
              <h3 className="text-sm font-semibold text-gray-700 mb-4 uppercase">PER AREA</h3>
              <div className="border rounded-lg p-4 bg-gray-50" style={{ minHeight: '350px' }}>
                {fieldPerAreaChartData.length === 0 ? (
                  <div className="flex items-center justify-center h-full text-gray-500">
                    No area data available for the selected filters. Please check if PER_AREA sheet has data for this month and product type.
                  </div>
                ) : (
                  <ResponsiveContainer width="100%" height={300}>
                    <LineChart data={fieldPerAreaChartData}>
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis dataKey="date" tick={{ fontSize: 10 }} angle={-45} textAnchor="end" height={60} />
                      <YAxis tick={{ fontSize: 10 }} />
                      <RechartsTooltip />
                      <Legend />
                      {areaNames.map((area, index) => (
                        <Line 
                          key={area} 
                          type="monotone" 
                          dataKey={area} 
                          stroke={getRandomColor(index)} 
                          strokeWidth={2} 
                        />
                      ))}
                    </LineChart>
                  </ResponsiveContainer>
                )}
              </div>
            </div>
          </>
        ) : activeTab === 'fieldCampaign' ? (
          <>
            <div className="bg-white rounded-lg p-6 shadow-sm border mb-6">
              <div className="grid grid-cols-4 gap-4">
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-2">Month</label>
                  <select value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)} className="w-full p-2 border border-gray-300 rounded text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500">
                    {availableMonths.map(m => <option key={m} value={m}>{m}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-2">Client</label>
                  <select value={selectedClient} onChange={(e) => setSelectedClient(e.target.value)} className="w-full p-2 border border-gray-300 rounded text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500">
                    {availableClients.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-2">Start Date</label>
                  <input type="date" value={fieldStartDate} onChange={(e) => setFieldStartDate(e.target.value)} className="w-full p-2 border border-gray-300 rounded text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500" />
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-2">End Date</label>
                  <input type="date" value={fieldEndDate} onChange={(e) => setFieldEndDate(e.target.value)} className="w-full p-2 border border-gray-300 rounded text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500" />
                </div>
              </div>
            </div>

            {fieldCampaignMetrics && (
              <div className="grid grid-cols-4 gap-6 mb-6">
                <div className="bg-white border rounded-lg p-6 shadow-sm">
                  <div className="text-sm font-semibold text-gray-600 mb-2">Total Portfolio</div>
                  <div className="text-3xl font-bold text-gray-900">{formatNumber(fieldCampaignMetrics.gross)}</div>
                  <div className="text-xs text-gray-500 mt-2">
                    BOM: {formatNumber(fieldCampaignMetrics.fieldBom)} + New Endo: {formatNumber(fieldCampaignMetrics.gross - fieldCampaignMetrics.fieldBom)}
                  </div>
                </div>
                <div className="bg-white border rounded-lg p-6 shadow-sm">
                  <div className="text-sm font-semibold text-gray-600 mb-2">Field BOM</div>
                  <div className="text-3xl font-bold text-gray-900">{formatNumber(fieldCampaignMetrics.fieldBom)}</div>
                  <div className="text-xs text-gray-500 mt-2">
                    Beginning of month accounts
                  </div>
                </div>
                <div className="bg-white border rounded-lg p-6 shadow-sm">
                  <div className="text-sm font-semibold text-gray-600 mb-2">VISITED</div>
                  <div className="text-3xl font-bold text-green-600">{formatNumber(fieldCampaignMetrics.totalVisited)}</div>
                  <div className="text-xs text-gray-500 mt-2">
                    Total accounts visited (MTD)
                  </div>
                </div>
                <div className="bg-white border rounded-lg p-6 shadow-sm">
                  <div className="text-sm font-semibold text-gray-600 mb-2">PENDING</div>
                  <div className="text-3xl font-bold text-orange-600">{formatNumber(fieldCampaignMetrics.totalPending)}</div>
                  <div className="text-xs text-gray-500 mt-2">
                    Gross: {formatNumber(fieldCampaignMetrics.gross)} - Visited: {formatNumber(fieldCampaignMetrics.totalVisited)}
                  </div>
                </div>
              </div>
            )}

            <div className="bg-white rounded-lg p-6 shadow-sm border mb-6">
              <h3 className="text-sm font-semibold text-gray-700 mb-4 uppercase">{selectedClient} - DAILY VISITATION</h3>
              <div className="border rounded-lg p-4 bg-gray-50" style={{ minHeight: '350px' }}>
                {fieldCampaignChartData.length === 0 ? (
                  <div className="flex items-center justify-center h-full text-gray-500">
                    No field data available for the selected client and filters.
                  </div>
                ) : (
                  <ResponsiveContainer width="100%" height={300}>
                    <LineChart data={fieldCampaignChartData}>
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis dataKey="date" tick={{ fontSize: 10 }} angle={-45} textAnchor="end" height={60} />
                      <YAxis tick={{ fontSize: 10 }} />
                      <RechartsTooltip />
                      <Legend />
                      <Line type="monotone" dataKey="visited" stroke="#3b82f6" strokeWidth={2} name="Visited">
                        <LabelList dataKey="visited" position="top" style={{ fontSize: '11px', fontWeight: 'bold', fill: '#3b82f6' }} />
                      </Line>
                    </LineChart>
                  </ResponsiveContainer>
                )}
              </div>
            </div>

            <div className="bg-white rounded-lg p-6 shadow-sm border">
              <h3 className="text-sm font-semibold text-gray-700 mb-4 uppercase">{selectedClient} - PER AREA</h3>
              <div className="border rounded-lg p-4 bg-gray-50" style={{ minHeight: '350px' }}>
                {fieldPerAreaByClientChartData.length === 0 ? (
                  <div className="flex items-center justify-center h-full text-gray-500">
                    No area data available for the selected client and filters.
                  </div>
                ) : (
                  <ResponsiveContainer width="100%" height={300}>
                    <LineChart data={fieldPerAreaByClientChartData}>
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis dataKey="date" tick={{ fontSize: 10 }} angle={-45} textAnchor="end" height={60} />
                      <YAxis tick={{ fontSize: 10 }} />
                      <RechartsTooltip />
                      <Legend />
                      {areaNamesForClient.map((area, index) => (
                        <Line 
                          key={area} 
                          type="monotone" 
                          dataKey={area} 
                          stroke={getRandomColor(index)} 
                          strokeWidth={2} 
                        />
                      ))}
                    </LineChart>
                  </ResponsiveContainer>
                )}
              </div>
            </div>
          </>
        ) : (
          <>
            {activeTab !== 'monthly' && (
              <div className="flex gap-4 mb-6">
                <div className="flex-1">
                  <label className="block text-xs font-medium text-gray-600 mb-1">Month</label>
                  <select value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)} className="w-full p-2 border rounded text-sm">
                    {availableMonths.map(m => <option key={m} value={m}>{m}</option>)}
                  </select>
                </div>
                {activeTab === 'overall' && (
                  <div className="flex-1">
                    <label className="block text-xs font-medium text-gray-600 mb-1">Product Type</label>
                    <select value={selectedProductType} onChange={(e) => setSelectedProductType(e.target.value)} className="w-full p-2 border rounded text-sm">
                      {availableProductTypes.map(t => <option key={t} value={t}>{t}</option>)}
                    </select>
                  </div>
                )}
                {activeTab === 'client' && (
                  <div className="flex-1">
                    <label className="block text-xs font-medium text-gray-600 mb-1">Client</label>
                    <select value={selectedClient} onChange={(e) => setSelectedClient(e.target.value)} className="w-full p-2 border rounded text-sm">
                      {availableClients.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>
                )}
                <div className="flex-1">
                  <label className="block text-xs font-medium text-gray-600 mb-1">Start Date</label>
                  <input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} className="w-full p-2 border rounded text-sm" />
                </div>
                <div className="flex-1">
                  <label className="block text-xs font-medium text-gray-600 mb-1">End Date</label>
                  <input type="date" value={endDate} onChange={(e) => setEndDate(e.target.value)} className="w-full p-2 border rounded text-sm" />
                </div>
              </div>
            )}

            {activeTab === 'monthly' && (
              <div className="bg-white rounded-lg p-6 shadow-md mb-6">
                <div className="flex justify-between items-center mb-4">
                  <h2 className="text-xl font-bold">📅 MTD Comparison Filters</h2>
                  <div className="text-sm bg-blue-50 px-3 py-1 rounded">Day {monthlyComparisonData.maxDay} MTD</div>
                </div>
                <div className="grid grid-cols-3 gap-4">
                  <MultiSelectDropdown
                    label="Select Months"
                    options={availableMonths}
                    value={selectedMonths}
                    onChange={setSelectedMonths}
                  />
                  <div>
                    <label className="block text-xs font-medium text-gray-600 mb-2">Product Type</label>
                    <select value={selectedComparisonProductType} onChange={(e) => setSelectedComparisonProductType(e.target.value)} className="w-full p-2 border border-gray-300 rounded text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500">
                      {availableProductTypes.map(t => <option key={t} value={t}>{t}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-gray-600 mb-2">View Mode</label>
                    <div className="flex gap-2">
                      <button onClick={() => setViewMode('count')} className={`flex-1 p-2 rounded transition-colors ${viewMode === 'count' ? 'bg-red-500 text-white' : 'bg-gray-100 hover:bg-gray-200'}`} title="Count">
                        <Hash className="h-5 w-5 mx-auto" />
                      </button>
                      <button onClick={() => setViewMode('ob')} className={`flex-1 p-2 rounded transition-colors ${viewMode === 'ob' ? 'bg-red-500 text-white' : 'bg-gray-100 hover:bg-gray-200'}`} title="Outstanding Balance">
                        <DollarSign className="h-5 w-5 mx-auto" />
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {(activeTab === 'overall' && overallMetrics) && (
              <div className="bg-white border rounded-lg p-6 mb-6">
                <div className="flex justify-between items-center mb-4">
                  <h2 className="text-lg font-bold">{selectedProductType} - SUMMARY</h2>
                  <div className="flex gap-2">
                    <button onClick={() => setViewMode('count')} className={`p-2 rounded ${viewMode === 'count' ? 'bg-red-500 text-white' : 'bg-gray-100'}`}><Hash className="h-5 w-5" /></button>
                    <button onClick={() => setViewMode('ob')} className={`p-2 rounded ${viewMode === 'ob' ? 'bg-red-500 text-white' : 'bg-gray-100'}`}><DollarSign className="h-5 w-5" /></button>
                  </div>
                </div>
                <div className="grid grid-cols-4 gap-6">
                  <div>
                    <div className="flex items-center gap-1 text-sm font-semibold mb-1">
                      BOM <InfoTooltip text="Beginning of Month - Starting portfolio value" />
                    </div>
                    <div className="text-2xl font-bold">{formatNumber(overallMetrics.bom)}</div>
                  </div>
                  <div>
                    <div className="flex items-center gap-1 text-sm font-semibold mb-1">
                      ACTIVE <InfoTooltip text="Current active portfolio value" />
                    </div>
                    <div className="text-2xl font-bold">{formatNumber(overallMetrics.active)}</div>
                  </div>
                  <div>
                    <div className="flex items-center gap-1 text-sm font-semibold mb-1">
                      NET GROWTH <InfoTooltip text="Percentage growth from BOM to Active" />
                    </div>
                    <div className="flex items-center gap-2">
                      <span className="text-2xl font-bold">{formatPercent(overallMetrics.portfolioGrowth)}</span>
                      <span className={`text-xl ${getGrowthIndicator(overallMetrics.portfolioGrowth).color}`}>
                        {getGrowthIndicator(overallMetrics.portfolioGrowth).icon}
                      </span>
                    </div>
                  </div>
                  <div>
                    <div className="flex items-center gap-1 text-sm font-semibold mb-1">
                      NET FLOW <InfoTooltip text="Ratio of endorsements to pullouts (higher is better)" />
                    </div>
                    <div className="flex items-center gap-2">
                      <span className="text-2xl font-bold">{formatNetFlow(overallMetrics.netFlowObj)}</span>
                      <span className={`text-xl ${getNetFlowIndicator(overallMetrics.netFlowObj).color}`}>
                        {getNetFlowIndicator(overallMetrics.netFlowObj).icon}
                      </span>
                    </div>
                  </div>
                </div>
                
                <div className="mt-6">
                  <h3 className="text-sm font-semibold mb-3">NET FLOW DAILY</h3>
                  <ResponsiveContainer width="100%" height={250}>
                    <LineChart data={getChartData}>
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis dataKey="date" tick={{ fontSize: 10 }} angle={-45} textAnchor="end" height={60} />
                      <YAxis tick={{ fontSize: 10 }} />
                      <RechartsTooltip />
                      <Line type="monotone" dataKey="netFlowRatio" stroke="#f97316" strokeWidth={2} />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              </div>
            )}

            {(activeTab === 'client' && clientMetrics) && (
              <div className="bg-white border rounded-lg p-6 mb-6">
                <div className="flex justify-between items-center mb-4">
                  <h2 className="text-lg font-bold">{selectedClient} - SUMMARY</h2>
                  <div className="flex gap-2">
                    <button onClick={() => setViewMode('count')} className={`p-2 rounded ${viewMode === 'count' ? 'bg-red-500 text-white' : 'bg-gray-100'}`}><Hash className="h-5 w-5" /></button>
                    <button onClick={() => setViewMode('ob')} className={`p-2 rounded ${viewMode === 'ob' ? 'bg-red-500 text-white' : 'bg-gray-100'}`}><DollarSign className="h-5 w-5" /></button>
                  </div>
                </div>
                <div className="grid grid-cols-4 gap-6">
                  <div>
                    <div className="text-sm font-semibold mb-1">BOM</div>
                    <div className="text-2xl font-bold">{formatNumber(clientMetrics.bom)}</div>
                  </div>
                  <div>
                    <div className="text-sm font-semibold mb-1">ACTIVE</div>
                    <div className="text-2xl font-bold">{formatNumber(clientMetrics.active)}</div>
                  </div>
                  <div>
                    <div className="text-sm font-semibold mb-1">NET GROWTH</div>
                    <div className="flex items-center gap-2">
                      <span className="text-2xl font-bold">{formatPercent(clientMetrics.portfolioGrowth)}</span>
                      <span className={`text-xl ${getGrowthIndicator(clientMetrics.portfolioGrowth).color}`}>
                        {getGrowthIndicator(clientMetrics.portfolioGrowth).icon}
                      </span>
                    </div>
                  </div>
                  <div>
                    <div className="text-sm font-semibold mb-1">NET FLOW</div>
                    <div className="flex items-center gap-2">
                      <span className="text-2xl font-bold">{formatNetFlow(clientMetrics.netFlowObj)}</span>
                      <span className={`text-xl ${getNetFlowIndicator(clientMetrics.netFlowObj).color}`}>
                        {getNetFlowIndicator(clientMetrics.netFlowObj).icon}
                      </span>
                    </div>
                  </div>
                </div>
                
                <div className="mt-6">
                  <h3 className="text-sm font-semibold mb-3">NET FLOW DAILY</h3>
                  <ResponsiveContainer width="100%" height={250}>
                    <LineChart data={getChartData}>
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis dataKey="date" tick={{ fontSize: 10 }} angle={-45} textAnchor="end" height={60} />
                      <YAxis tick={{ fontSize: 10 }} />
                      <RechartsTooltip />
                      <Line type="monotone" dataKey="netFlowRatio" stroke="#f97316" strokeWidth={2} />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              </div>
            )}

            {activeTab !== 'monthly' && (
              <>
                <div className="bg-white rounded-lg p-6 shadow-md mb-6">
                  <h2 className="text-xl font-bold mb-4">Daily Movement</h2>
                  <div className="space-y-6">
                    <div>
                      <h3 className="text-sm font-semibold mb-3">Daily Endorsements vs Pullouts</h3>
                      <ResponsiveContainer width="100%" height={250}>
                        <LineChart data={getChartData}>
                          <CartesianGrid strokeDasharray="3 3" />
                          <XAxis dataKey="date" tick={{ fontSize: 10 }} />
                          <YAxis />
                          <RechartsTooltip />
                          <Legend />
                          <Line type="monotone" dataKey="endorsements" stroke="#10b981" strokeWidth={2} />
                          <Line type="monotone" dataKey="pullouts" stroke="#ef4444" strokeWidth={2} />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                    <div>
                      <h3 className="text-sm font-semibold mb-3">Daily Net Growth</h3>
                      <ResponsiveContainer width="100%" height={250}>
                        <LineChart data={getChartData}>
                          <CartesianGrid strokeDasharray="3 3" />
                          <XAxis dataKey="date" tick={{ fontSize: 10 }} />
                          <YAxis />
                          <RechartsTooltip />
                          <Legend />
                          <Line type="monotone" dataKey="portfolioGrowth" stroke="#8b5cf6" strokeWidth={2} />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                </div>

                <div className="bg-white rounded-lg p-6 shadow-md">
                  <h2 className="text-xl font-bold mb-4">MTD Trends</h2>
                  <div className="space-y-6">
                    <div>
                      <h3 className="text-sm font-semibold mb-3">Cumulative Endorsements vs Pullouts</h3>
                      <ResponsiveContainer width="100%" height={250}>
                        <LineChart data={getMTDData}>
                          <CartesianGrid strokeDasharray="3 3" />
                          <XAxis dataKey="date" tick={{ fontSize: 10 }} />
                          <YAxis />
                          <RechartsTooltip />
                          <Legend />
                          <Line type="monotone" dataKey="cumulativeEndorsements" stroke="#10b981" strokeWidth={2} />
                          <Line type="monotone" dataKey="cumulativePullouts" stroke="#ef4444" strokeWidth={2} />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                    <div>
                      <h3 className="text-sm font-semibold mb-3">MTD Net Flow Ratio</h3>
                      <ResponsiveContainer width="100%" height={250}>
                        <LineChart data={getMTDData}>
                          <CartesianGrid strokeDasharray="3 3" />
                          <XAxis dataKey="date" tick={{ fontSize: 10 }} />
                          <YAxis />
                          <RechartsTooltip />
                          <Legend />
                          <Line type="monotone" dataKey="mtdNetFlowRatio" stroke="#3b82f6" strokeWidth={2} />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                    <div>
                      <h3 className="text-sm font-semibold mb-3">MTD Net Growth</h3>
                      <ResponsiveContainer width="100%" height={250}>
                        <LineChart data={getMTDData}>
                          <CartesianGrid strokeDasharray="3 3" />
                          <XAxis dataKey="date" tick={{ fontSize: 10 }} />
                          <YAxis />
                          <RechartsTooltip />
                          <Legend />
                          <Line type="monotone" dataKey="mtdPortfolioGrowth" stroke="#8b5cf6" strokeWidth={2} />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                </div>
              </>
            )}

            {activeTab === 'monthly' && (
              <>
                <div className="bg-white rounded-lg p-6 shadow-md mb-6">
                  <h2 className="text-xl font-bold mb-4">📊 MTD Performance Overview</h2>
                  <div className="grid grid-cols-3 gap-4">
                    {monthlyComparisonData.monthlyMetrics.map(m => (
                      <div key={m.month} className="border rounded-lg p-4">
                        <h3 className="text-lg font-bold text-indigo-600 mb-3">{m.month}</h3>
                        <div className="space-y-2">
                          <div className="flex justify-between"><span className="text-sm">BOM:</span><span className="font-semibold">{formatNumber(m.bom)}</span></div>
                          <div className="flex justify-between"><span className="text-sm">Active:</span><span className="font-semibold">{formatNumber(m.active)}</span></div>
                          <div className="flex justify-between items-center">
                            <span className="text-sm">Net Growth:</span>
                            <div className="flex items-center gap-1">
                              <span className="font-semibold">{formatPercent(m.portfolioGrowth)}</span>
                              <span className={`${getGrowthIndicator(m.portfolioGrowth).color}`}>{getGrowthIndicator(m.portfolioGrowth).icon}</span>
                            </div>
                          </div>
                          <div className="flex justify-between items-center">
                            <span className="text-sm">Net Flow:</span>
                            <div className="flex items-center gap-1">
                              <span className="font-semibold">{formatNetFlow(m.netFlowObj)}</span>
                              <span className={`${getNetFlowIndicator(m.netFlowObj).color}`}>{getNetFlowIndicator(m.netFlowObj).icon}</span>
                            </div>
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>

                <div className="bg-white rounded-lg p-6 shadow-md mb-6">
                  <h2 className="text-xl font-bold mb-4">📈 MTD Endorsements vs Pullouts</h2>
                  <ResponsiveContainer width="100%" height={300}>
                    <BarChart data={monthlyComparisonData.monthlyMetrics}>
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis dataKey="month" />
                      <YAxis />
                      <RechartsTooltip />
                      <Legend />
                      <Bar dataKey="totalEndorsements" fill="#10b981" />
                      <Bar dataKey="totalPullouts" fill="#ef4444" />
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div className="bg-white rounded-lg p-6 shadow-md">
                  <div className="flex justify-between items-center mb-4">
                    <h2 className="text-xl font-bold">🏢 Client-Level MTD Comparison</h2>
                    <div className="flex gap-2">
                      <button onClick={() => setClientRankingView('all')} className={`px-3 py-1 rounded text-sm ${clientRankingView === 'all' ? 'bg-indigo-500 text-white' : 'bg-gray-200'}`}>All</button>
                      <button onClick={() => setClientRankingView('top5')} className={`px-3 py-1 rounded text-sm ${clientRankingView === 'top5' ? 'bg-green-500 text-white' : 'bg-gray-200'}`}>Top 5</button>
                      <button onClick={() => setClientRankingView('bottom5')} className={`px-3 py-1 rounded text-sm ${clientRankingView === 'bottom5' ? 'bg-red-500 text-white' : 'bg-gray-200'}`}>Bottom 5</button>
                    </div>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full border-collapse">
                      <thead>
                        <tr className="bg-gray-100">
                          <th className="border p-2 text-left text-sm">Client</th>
                          <th className="border p-2 text-right text-sm">BOM</th>
                          <th className="border p-2 text-right text-sm">Active</th>
                          <th className="border p-2 text-right text-sm">Net Growth %</th>
                          <th className="border p-2 text-right text-sm">Net Flow</th>
                        </tr>
                      </thead>
                      <tbody>
                        {monthlyComparisonData.clientComparison.map(c => (
                          <tr key={c.name} className="hover:bg-gray-50">
                            <td className="border p-2 text-sm">{c.name}</td>
                            <td className="border p-2 text-right text-sm">{formatNumber(c.bom)}</td>
                            <td className="border p-2 text-right text-sm">{formatNumber(c.active)}</td>
                            <td className="border p-2 text-right text-sm">
                              <div className="flex items-center justify-end gap-1">
                                <span>{formatPercent(c.portfolioGrowth)}</span>
                                <span className={`${getGrowthIndicator(c.portfolioGrowth).color}`}>{getGrowthIndicator(c.portfolioGrowth).icon}</span>
                              </div>
                            </td>
                            <td className="border p-2 text-right text-sm">
                              <div className="flex items-center justify-end gap-1">
                                <span>{formatNetFlow(c.netFlowObj)}</span>
                                <span className={`${getNetFlowIndicator(c.netFlowObj).color}`}>{getNetFlowIndicator(c.netFlowObj).icon}</span>
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </>
            )}
          </>
        )}
      </div>
    </div>
  );
}