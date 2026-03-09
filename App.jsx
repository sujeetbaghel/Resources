import React, { useState, useMemo, useEffect, useCallback } from "react";
import { 
  Database, 
  LogOut, 
  RefreshCw, 
  ChevronDown, 
  ChevronRight,
  Search,
  X,
  Check,
  Settings,
  Table,
  Columns,
  Eye,
  Download,
  AlertCircle,
  Loader2,
  ChevronsUpDown,
  Filter,
  Play,
  Terminal,
  Info
} from "lucide-react";

// Debounce hook for search inputs
const useDebounce = (value, delay = 300) => {
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
};

const App = () => {
  const [token, setToken] = useState(null);
  const [environmentUrl, setEnvironmentUrl] = useState('');
  const [isAuthenticating, setIsAuthenticating] = useState(false);
  const [apiResponse, setApiResponse] = useState(null);
  const [isLoading, setIsLoading] = useState(false);
  
  // Entity explorer state
  const [entities, setEntities] = useState([]);
  const [selectedEntity, setSelectedEntity] = useState('');
  const [recordCount, setRecordCount] = useState('10');
  const [customCount, setCustomCount] = useState('');
  const [showCustomCount, setShowCustomCount] = useState(false);
  const [expandedRows, setExpandedRows] = useState({});
  const [entityMetadata, setEntityMetadata] = useState({});
  const [fetchError, setFetchError] = useState(null);
  const [selectedFields, setSelectedFields] = useState([]);
  const [customField, setCustomField] = useState('');
  const [showFieldSelector, setShowFieldSelector] = useState(false);
  const [entityFields, setEntityFields] = useState([]);
  
  // Search states
  const [entitySearch, setEntitySearch] = useState('');
  const [fieldSearch, setFieldSearch] = useState('');
  const debouncedEntitySearch = useDebounce(entitySearch);
  const debouncedFieldSearch = useDebounce(fieldSearch);
  
  // View mode for response
  const [viewMode, setViewMode] = useState('pretty'); // 'pretty' or 'raw'

  useEffect(() => {
    const hash = window.location.hash;
    if (hash && hash.includes('access_token')) {
      const params = new URLSearchParams(hash.substring(1));
      const accessToken = params.get('access_token');
      if (accessToken) {
        setToken(accessToken);
        const savedEnv = sessionStorage.getItem('dataverse_env');
        if (savedEnv) {
          setEnvironmentUrl(savedEnv);
        }
        window.history.replaceState(null, '', '/');
      }
    }
  }, []);

  const apiBaseUrl = useMemo(() => {
    return environmentUrl ? `${environmentUrl.replace(/\/$/, '')}/api/data/v9.2` : '';
  }, [environmentUrl]);

  const authHeaders = useMemo(() => {
    if (!token) return null;
    return {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      'Content-Type': 'application/json',
      'Prefer': 'odata.include-annotations="*"'
    };
  }, [token]);

  // Fetch entities
  useEffect(() => {
    if (token && apiBaseUrl) {
      fetchEntities();
    }
  }, [token, apiBaseUrl]);

  const fetchEntities = async () => {
    setIsLoading(true);
    setFetchError(null);
    try {
      const response = await fetch(
        `${apiBaseUrl}/EntityDefinitions`,
        { headers: authHeaders }
      );

      if (!response.ok) throw new Error(`Failed to fetch entities: ${response.status}`);
      
      const data = await response.json();
      
      const filteredEntities = data.value
        .filter(entity => !entity.IsPrivate && !entity.IsLogicalEntity)
        .map(entity => ({
          LogicalName: entity.LogicalName,
          DisplayName: entity.DisplayName?.UserLocalizedLabel?.Label || entity.LogicalName,
          EntitySetName: entity.EntitySetName || entity.LogicalName,
          PrimaryIdAttribute: entity.PrimaryIdAttribute,
          PrimaryNameAttribute: entity.PrimaryNameAttribute
        }))
        .sort((a, b) => a.DisplayName.localeCompare(b.DisplayName));
      
      setEntities(filteredEntities);
      
      const metadataMap = {};
      filteredEntities.forEach(entity => {
        metadataMap[entity.EntitySetName] = {
          primaryId: entity.PrimaryIdAttribute,
          primaryName: entity.PrimaryNameAttribute,
          displayName: entity.DisplayName,
          logicalName: entity.LogicalName
        };
      });
      setEntityMetadata(metadataMap);
      
    } catch (error) {
      console.error('Error fetching entities:', error);
      setFetchError(error.message);
    } finally {
      setIsLoading(false);
    }
  };

  const fetchEntityFields = async (entitySetName) => {
    try {
      const metadata = entityMetadata[entitySetName];
      if (!metadata) return;
      
      const response = await fetch(
        `${apiBaseUrl}/EntityDefinitions(LogicalName='${metadata.logicalName}')/Attributes?$select=LogicalName,DisplayName,AttributeType&$filter=IsLogical eq false and AttributeOf eq null`,
        { headers: authHeaders }
      );

      if (!response.ok) throw new Error(`Failed to fetch fields: ${response.status}`);
      
      const data = await response.json();
      const fields = data.value
        .map(field => ({
          logicalName: field.LogicalName,
          displayName: field.DisplayName?.UserLocalizedLabel?.Label || field.LogicalName,
          type: field.AttributeType
        }))
        .sort((a, b) => a.displayName.localeCompare(b.displayName));
      
      setEntityFields(fields);
    } catch (error) {
      console.error('Error fetching fields:', error);
    }
  };

  const fetchEntityData = async () => {
    if (!selectedEntity) {
      alert('Please select an entity');
      return;
    }

    setIsLoading(true);
    setApiResponse(null);
    setFetchError(null);
    setExpandedRows({});

    try {
      const count = showCustomCount ? customCount : recordCount;
      const entitySet = selectedEntity;
      
      let selectFields = '';
      if (selectedFields.length > 0) {
        selectFields = `&$select=${selectedFields.join(',')}`;
      }
      
      const url = `${apiBaseUrl}/${entitySet}?$top=${count}${selectFields}`;
      
      const response = await fetch(url, { headers: authHeaders });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`HTTP ${response.status}: ${errorText}`);
      }

      const data = await response.json();
      setApiResponse(data);
      
    } catch (error) {
      console.error('Error fetching entity data:', error);
      setFetchError(error.message);
    } finally {
      setIsLoading(false);
    }
  };

  const handleEntityChange = async (e) => {
    const entitySet = e.target.value;
    setSelectedEntity(entitySet);
    setSelectedFields([]);
    setEntityFields([]);
    setFieldSearch('');
    
    if (entitySet) {
      await fetchEntityFields(entitySet);
    }
  };

  const handleLogin = () => {
    if (!environmentUrl) {
      alert('Please enter your Dataverse environment URL');
      return;
    }

    sessionStorage.setItem('dataverse_env', environmentUrl);
    setIsAuthenticating(true);

    const orgUrl = environmentUrl.replace(/\/$/, '');
    const clientId = '51f81489-12ee-4a9e-aaae-a2591f45987d';
    const redirectUri = encodeURIComponent('http://localhost:5173');
    const scope = encodeURIComponent(`${orgUrl}/.default`);

    const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?
      client_id=${clientId}
      &response_type=token
      &redirect_uri=${redirectUri}
      &scope=${scope}
      &response_mode=fragment`.replace(/\s+/g, '');

    window.location.href = authUrl;
  };

  const logout = () => {
    setToken(null);
    setApiResponse(null);
    setEntities([]);
    setSelectedEntity('');
    setEntityFields([]);
    sessionStorage.removeItem('dataverse_env');
  };

  const toggleRowExpand = (index) => {
    setExpandedRows(prev => ({
      ...prev,
      [index]: !prev[index]
    }));
  };

  const toggleFieldSelection = (field) => {
    setSelectedFields(prev => {
      if (prev.includes(field)) {
        return prev.filter(f => f !== field);
      } else {
        return [...prev, field];
      }
    });
  };

  const addCustomField = () => {
    if (customField && !selectedFields.includes(customField)) {
      setSelectedFields([...selectedFields, customField]);
      setCustomField('');
    }
  };

  const clearAllFields = () => {
    setSelectedFields([]);
  };

  const exportData = () => {
    if (!apiResponse) return;
    
    const dataStr = JSON.stringify(apiResponse, null, 2);
    const dataUri = 'data:application/json;charset=utf-8,'+ encodeURIComponent(dataStr);
    const exportFileDefaultName = `dataverse-${selectedEntity}-${new Date().toISOString()}.json`;
    
    const linkElement = document.createElement('a');
    linkElement.setAttribute('href', dataUri);
    linkElement.setAttribute('download', exportFileDefaultName);
    linkElement.click();
  };

  // Filter entities based on search
  const filteredEntities = useMemo(() => {
    if (!debouncedEntitySearch) return entities;
    const searchLower = debouncedEntitySearch.toLowerCase();
    return entities.filter(entity => 
      entity.DisplayName.toLowerCase().includes(searchLower) ||
      entity.EntitySetName.toLowerCase().includes(searchLower) ||
      entity.LogicalName.toLowerCase().includes(searchLower)
    );
  }, [entities, debouncedEntitySearch]);

  // Filter fields based on search
  const filteredFields = useMemo(() => {
    if (!debouncedFieldSearch) return entityFields;
    const searchLower = debouncedFieldSearch.toLowerCase();
    return entityFields.filter(field => 
      field.displayName.toLowerCase().includes(searchLower) ||
      field.logicalName.toLowerCase().includes(searchLower) ||
      field.type.toLowerCase().includes(searchLower)
    );
  }, [entityFields, debouncedFieldSearch]);

  const renderValue = (value) => {
    if (value === null || value === undefined) {
      return <span className="text-gray-400 italic">null</span>;
    }
    if (typeof value === 'object') {
      return <span className="text-blue-500">[Object]</span>;
    }
    if (typeof value === 'boolean') {
      return <span className="text-purple-500">{value.toString()}</span>;
    }
    if (typeof value === 'string' && value.length > 100) {
      return <span title={value} className="text-gray-700">{value.substring(0, 100)}...</span>;
    }
    if (typeof value === 'number') {
      return <span className="text-green-600">{value}</span>;
    }
    return <span className="text-gray-800">{value.toString()}</span>;
  };

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <header className="bg-white border-b border-gray-200 sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center space-x-3">
              <div className="bg-blue-600 p-2 rounded-lg">
                <Database className="w-6 h-6 text-white" />
              </div>
              <h1 className="text-xl font-semibold text-gray-900">Dataverse Explorer</h1>
              {token && (
                <span className="ml-2 px-2 py-1 bg-green-100 text-green-700 text-xs font-medium rounded-full">
                  Connected
                </span>
              )}
            </div>
            {token && (
              <button
                onClick={logout}
                className="inline-flex items-center px-3 py-2 border border-gray-300 shadow-sm text-sm leading-4 font-medium rounded-md text-gray-700 bg-white hover:bg-gray-50 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 transition-colors"
              >
                <LogOut className="w-4 h-4 mr-2" />
                Disconnect
              </button>
            )}
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {!token ? (
          // Login Card
          <div className="max-w-md mx-auto">
            <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-8">
              <div className="text-center mb-8">
                <div className="bg-blue-50 p-3 rounded-full w-16 h-16 mx-auto mb-4 flex items-center justify-center">
                  <Database className="w-8 h-8 text-blue-600" />
                </div>
                <h2 className="text-2xl font-semibold text-gray-900">Connect to Dataverse</h2>
                <p className="mt-2 text-sm text-gray-600">
                  Enter your environment URL to start exploring
                </p>
              </div>
              
              <div className="space-y-4">
                <div>
                  <label htmlFor="envUrl" className="block text-sm font-medium text-gray-700 mb-1">
                    Environment URL
                  </label>
                  <input
                    id="envUrl"
                    type="text"
                    value={environmentUrl}
                    onChange={(e) => setEnvironmentUrl(e.target.value)}
                    placeholder="https://yourorg.crm.dynamics.com"
                    className="w-full px-3 py-2 border border-gray-300 rounded-lg shadow-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-colors"
                  />
                  <p className="mt-1 text-xs text-gray-500">
                    Example: https://myorg.crm.dynamics.com
                  </p>
                </div>
                
                <button
                  onClick={handleLogin}
                  disabled={isAuthenticating || !environmentUrl}
                  className="w-full inline-flex justify-center items-center px-4 py-2 border border-transparent shadow-sm text-sm font-medium rounded-lg text-white bg-blue-600 hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
                >
                  {isAuthenticating ? (
                    <>
                      <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                      Redirecting...
                    </>
                  ) : (
                    'Connect'
                  )}
                </button>
              </div>
            </div>
          </div>
        ) : (
          // Main Explorer Interface
          <div className="space-y-6">
            {/* Control Panel */}
            <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-6">
              <div className="flex items-center justify-between mb-6">
                <h2 className="text-lg font-semibold text-gray-900 flex items-center">
                  <Settings className="w-5 h-5 mr-2 text-gray-500" />
                  Query Builder
                </h2>
                <div className="text-sm text-gray-500">
                  {environmentUrl}
                </div>
              </div>

              <div className="space-y-4">
                {/* Entity Selector with Search */}
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Select Entity
                  </label>
                  <div className="relative">
                    <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                      <Search className="h-4 w-4 text-gray-400" />
                    </div>
                    <input
                      type="text"
                      value={entitySearch}
                      onChange={(e) => setEntitySearch(e.target.value)}
                      placeholder="Search entities..."
                      className="block w-full pl-10 pr-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                    />
                  </div>
                  
                  <div className="mt-2 relative">
                    <select
                      value={selectedEntity}
                      onChange={handleEntityChange}
                      className="block w-full pl-3 pr-10 py-2 text-sm border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 appearance-none bg-white"
                      style={{ backgroundImage: 'none' }}
                    >
                      <option value="">-- Choose an entity --</option>
                      {filteredEntities.map(entity => (
                        <option key={entity.EntitySetName} value={entity.EntitySetName}>
                          {entity.DisplayName} ({entity.EntitySetName})
                        </option>
                      ))}
                    </select>
                    <div className="absolute inset-y-0 right-0 flex items-center pr-2 pointer-events-none">
                      <ChevronsUpDown className="h-4 w-4 text-gray-400" />
                    </div>
                  </div>
                  
                  {filteredEntities.length === 0 && entitySearch && (
                    <p className="mt-1 text-xs text-gray-500">No entities match your search</p>
                  )}
                </div>

                {/* Record Count and Controls */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">
                      Records to Fetch
                    </label>
                    <div className="flex space-x-2">
                      <select
                        value={recordCount}
                        onChange={(e) => {
                          setRecordCount(e.target.value);
                          setShowCustomCount(e.target.value === 'custom');
                        }}
                        className="block w-32 pl-3 pr-10 py-2 text-sm border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                      >
                        <option value="10">10</option>
                        <option value="25">25</option>
                        <option value="50">50</option>
                        <option value="100">100</option>
                        <option value="250">250</option>
                        <option value="500">500</option>
                        <option value="1000">1000</option>
                        <option value="custom">Custom</option>
                      </select>
                      
                      {showCustomCount && (
                        <input
                          type="number"
                          value={customCount}
                          onChange={(e) => setCustomCount(e.target.value)}
                          placeholder="Count"
                          min="1"
                          max="5000"
                          className="block w-32 px-3 py-2 text-sm border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                        />
                      )}
                    </div>
                  </div>

                  <div className="flex items-end space-x-2">
                    <button
                      onClick={() => setShowFieldSelector(!showFieldSelector)}
                      disabled={!selectedEntity}
                      className={`inline-flex items-center px-4 py-2 border text-sm font-medium rounded-lg transition-colors ${
                        selectedEntity
                          ? 'border-gray-300 bg-white text-gray-700 hover:bg-gray-50'
                          : 'border-gray-200 bg-gray-50 text-gray-400 cursor-not-allowed'
                      }`}
                    >
                      <Columns className="w-4 h-4 mr-2" />
                      {showFieldSelector ? 'Hide Fields' : 'Select Fields'}
                    </button>

                    <button
                      onClick={fetchEntityData}
                      disabled={!selectedEntity || isLoading}
                      className={`inline-flex items-center px-6 py-2 border border-transparent text-sm font-medium rounded-lg shadow-sm text-white transition-colors ${
                        selectedEntity && !isLoading
                          ? 'bg-green-600 hover:bg-green-700'
                          : 'bg-gray-400 cursor-not-allowed'
                      }`}
                    >
                      {isLoading ? (
                        <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                      ) : (
                        <Play className="w-4 h-4 mr-2" />
                      )}
                      {isLoading ? 'Fetching...' : 'Fetch Data'}
                    </button>

                    <button
                      onClick={fetchEntities}
                      disabled={isLoading}
                      className="inline-flex items-center p-2 border border-gray-300 rounded-lg text-gray-700 bg-white hover:bg-gray-50 transition-colors disabled:opacity-50"
                      title="Refresh entities"
                    >
                      <RefreshCw className={`w-4 h-4 ${isLoading ? 'animate-spin' : ''}`} />
                    </button>
                  </div>
                </div>

                {/* Field Selector Panel */}
                {showFieldSelector && entityFields.length > 0 && (
                  <div className="mt-4 p-4 bg-gray-50 rounded-lg border border-gray-200">
                    <div className="flex items-center justify-between mb-3">
                      <h3 className="text-sm font-medium text-gray-700 flex items-center">
                        <Filter className="w-4 h-4 mr-1" />
                        Select Fields to Include
                      </h3>
                      <button
                        onClick={clearAllFields}
                        className="text-xs text-gray-500 hover:text-gray-700"
                      >
                        Clear all
                      </button>
                    </div>

                    {/* Field Search */}
                    <div className="relative mb-3">
                      <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                        <Search className="h-3 w-3 text-gray-400" />
                      </div>
                      <input
                        type="text"
                        value={fieldSearch}
                        onChange={(e) => setFieldSearch(e.target.value)}
                        placeholder="Search fields..."
                        className="block w-full pl-9 pr-3 py-1.5 text-xs border border-gray-300 rounded-md focus:ring-1 focus:ring-blue-500 focus:border-blue-500"
                      />
                    </div>

                    {/* Fields Grid */}
                    <div className="max-h-48 overflow-y-auto border border-gray-200 rounded-md bg-white">
                      {filteredFields.map(field => (
                        <label
                          key={field.logicalName}
                          className="flex items-center px-3 py-2 hover:bg-gray-50 cursor-pointer border-b border-gray-100 last:border-b-0"
                        >
                          <input
                            type="checkbox"
                            checked={selectedFields.includes(field.logicalName)}
                            onChange={() => toggleFieldSelection(field.logicalName)}
                            className="h-3.5 w-3.5 text-blue-600 rounded border-gray-300 focus:ring-blue-500"
                          />
                          <span className="ml-2 text-xs">
                            <span className="font-medium text-gray-700">{field.displayName}</span>
                            <span className="ml-1 text-gray-400 text-[10px]">
                              ({field.logicalName} - {field.type})
                            </span>
                          </span>
                        </label>
                      ))}
                    </div>

                    {/* Custom Field Input */}
                    <div className="mt-3 flex space-x-2">
                      <input
                        type="text"
                        value={customField}
                        onChange={(e) => setCustomField(e.target.value)}
                        placeholder="Add custom field name"
                        className="block flex-1 px-3 py-1.5 text-xs border border-gray-300 rounded-md focus:ring-1 focus:ring-blue-500 focus:border-blue-500"
                      />
                      <button
                        onClick={addCustomField}
                        className="px-3 py-1.5 text-xs bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors"
                      >
                        Add
                      </button>
                    </div>

                    {selectedFields.length > 0 && (
                      <div className="mt-2 text-xs text-gray-600">
                        Selected: {selectedFields.length} field{selectedFields.length !== 1 ? 's' : ''}
                      </div>
                    )}
                  </div>
                )}
              </div>
            </div>

            {/* Error Display */}
            {fetchError && (
              <div className="bg-red-50 border border-red-200 rounded-xl p-4">
                <div className="flex">
                  <AlertCircle className="h-5 w-5 text-red-400 mr-3" />
                  <div className="flex-1">
                    <h3 className="text-sm font-medium text-red-800">Error</h3>
                    <p className="text-sm text-red-700 mt-1">{fetchError}</p>
                  </div>
                  <button
                    onClick={() => setFetchError(null)}
                    className="text-red-400 hover:text-red-500"
                  >
                    <X className="h-5 w-5" />
                  </button>
                </div>
              </div>
            )}

            {/* Results Section */}
            {apiResponse && (
              <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
                {/* Results Header */}
                <div className="border-b border-gray-200 px-6 py-4 bg-gray-50">
                  <div className="flex items-center justify-between">
                    <div className="flex items-center space-x-4">
                      <div className="flex items-center space-x-2">
                        <Table className="h-5 w-5 text-gray-500" />
                        <h3 className="text-base font-medium text-gray-900">
                          {selectedEntity}
                        </h3>
                      </div>
                      <span className="px-2.5 py-0.5 bg-blue-100 text-blue-800 text-xs font-medium rounded-full">
                        {apiResponse.value?.length || 0} records
                      </span>
                    </div>
                    
                    <div className="flex items-center space-x-3">
                      {/* View Toggle */}
                      <div className="flex items-center border border-gray-200 rounded-lg overflow-hidden">
                        <button
                          onClick={() => setViewMode('pretty')}
                          className={`px-3 py-1.5 text-xs font-medium transition-colors ${
                            viewMode === 'pretty'
                              ? 'bg-blue-600 text-white'
                              : 'bg-white text-gray-700 hover:bg-gray-50'
                          }`}
                        >
                          <Eye className="w-3.5 h-3.5 inline mr-1" />
                          Pretty
                        </button>
                        <button
                          onClick={() => setViewMode('raw')}
                          className={`px-3 py-1.5 text-xs font-medium transition-colors ${
                            viewMode === 'raw'
                              ? 'bg-blue-600 text-white'
                              : 'bg-white text-gray-700 hover:bg-gray-50'
                          }`}
                        >
                          <Terminal className="w-3.5 h-3.5 inline mr-1" />
                          Raw
                        </button>
                      </div>

                      {/* Export Button */}
                      <button
                        onClick={exportData}
                        className="inline-flex items-center px-3 py-1.5 border border-gray-300 text-xs font-medium rounded-lg text-gray-700 bg-white hover:bg-gray-50 transition-colors"
                      >
                        <Download className="w-3.5 h-3.5 mr-1" />
                        Export
                      </button>
                    </div>
                  </div>
                </div>

                {/* Results Content */}
                <div className="p-6">
                  {viewMode === 'pretty' ? (
                    <div className="space-y-3 max-h-[600px] overflow-y-auto">
                      {apiResponse.value && apiResponse.value.length > 0 ? (
                        apiResponse.value.map((record, index) => {
                          const isExpanded = expandedRows[index];
                          const primaryName = entityMetadata[selectedEntity]?.primaryName;
                          const primaryValue = primaryName ? record[primaryName] : null;
                          
                          return (
                            <div
                              key={index}
                              className="border border-gray-200 rounded-lg overflow-hidden hover:border-blue-200 transition-colors"
                            >
                              <div
                                onClick={() => toggleRowExpand(index)}
                                className="flex items-center justify-between px-4 py-3 bg-gray-50 cursor-pointer hover:bg-gray-100 transition-colors"
                              >
                                <div className="flex items-center space-x-3">
                                  {isExpanded ? (
                                    <ChevronDown className="w-4 h-4 text-gray-500" />
                                  ) : (
                                    <ChevronRight className="w-4 h-4 text-gray-500" />
                                  )}
                                  <span className="text-sm font-medium text-gray-900">
                                    Record {index + 1}
                                  </span>
                                  {primaryValue && (
                                    <span className="text-xs text-gray-600 bg-white px-2 py-0.5 rounded-full border border-gray-200">
                                      {primaryValue}
                                    </span>
                                  )}
                                </div>
                                <span className="text-xs text-gray-500">
                                  {Object.keys(record).length} fields
                                </span>
                              </div>
                              
                              {isExpanded && (
                                <div className="px-4 py-3 bg-white border-t border-gray-200">
                                  <div className="grid grid-cols-1 gap-1">
                                    {Object.entries(record).map(([key, value]) => (
                                      <div key={key} className="flex text-sm py-1.5 border-b border-gray-50 last:border-0">
                                        <span className="w-1/3 font-medium text-gray-600 text-xs uppercase tracking-wider">
                                          {key}:
                                        </span>
                                        <span className="w-2/3 text-sm font-mono">
                                          {renderValue(value)}
                                        </span>
                                      </div>
                                    ))}
                                  </div>
                                </div>
                              )}
                            </div>
                          );
                        })
                      ) : (
                        <div className="text-center py-12">
                          <Info className="mx-auto h-12 w-12 text-gray-400" />
                          <h3 className="mt-2 text-sm font-medium text-gray-900">No records found</h3>
                          <p className="mt-1 text-sm text-gray-500">
                            This entity has no records or the query returned no results.
                          </p>
                        </div>
                      )}
                    </div>
                  ) : (
                    <pre className="bg-gray-900 text-gray-100 p-4 rounded-lg overflow-auto max-h-[600px] text-xs font-mono">
                      {JSON.stringify(apiResponse, null, 2)}
                    </pre>
                  )}
                </div>

                {/* Footer */}
                <div className="border-t border-gray-200 px-6 py-3 bg-gray-50">
                  <div className="flex items-center justify-between text-xs text-gray-500">
                    <span>
                      <strong>Entity:</strong> {selectedEntity}
                    </span>
                    <span>
                      <strong>Fields:</strong> {selectedFields.length || 'All'}
                    </span>
                    <span>
                      <strong>Records:</strong> {apiResponse.value?.length || 0}
                    </span>
                    {apiResponse['@odata.nextLink'] && (
                      <span className="text-yellow-600 flex items-center">
                        <Info className="w-3 h-3 mr-1" />
                        More records available
                      </span>
                    )}
                  </div>
                </div>
              </div>
            )}
          </div>
        )}
      </main>
    </div>
  );
};

export default App;
