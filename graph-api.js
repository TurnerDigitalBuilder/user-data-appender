// Microsoft Graph API Integration
const GraphAPI = (function() {
  'use strict';
  
  // Private variables
  let accessToken = '';
  let startingEmail = '';
  let maxUsers = 100;
  let licensedEmails = new Set(); // Stores full email addresses from CSV
  let licensedUsernames = new Set(); // Stores email usernames for matching
  let aiEngagementMap = new Map(); // Map of email usernames to daily messages
  let orgData = [];
  let filteredOrgData = [];
  let hierarchyData = null;
  let filterNoDepartment = true; // Default to filtering out users with no department
  let filterInterns = true; // Default to filtering out interns
  let filterNoTitle = true; // Default to filtering out users with no title
  
  // API configuration
  const graphBaseUrl = 'https://graph.microsoft.com/v1.0';
  const orgDomain = 'tcco.com';
  
  // Public API
  return {
    // Initialize the module
    initialize: function() {
      // Set up event listeners for CSV upload
      const csvInput = document.getElementById('csvFile');
      if (csvInput) {
        csvInput.addEventListener('change', this.handleCsvUpload.bind(this));
      }
      
      // Set up department filter checkbox listener
      const filterCheckbox = document.getElementById('filterNoDept');
      if (filterCheckbox) {
        filterCheckbox.checked = filterNoDepartment; // Set default state
        filterCheckbox.addEventListener('change', this.handleFilterChange.bind(this));
      }
      
      // Set up interns filter checkbox listener
      const filterInternsCheckbox = document.getElementById('filterInterns');
      if (filterInternsCheckbox) {
        filterInternsCheckbox.checked = filterInterns; // Set default state
        filterInternsCheckbox.addEventListener('change', this.handleFilterChange.bind(this));
      }

      // Set up title filter checkbox listener
      const filterNoTitleCheckbox = document.getElementById('filterNoTitle');
      if (filterNoTitleCheckbox) {
        filterNoTitleCheckbox.checked = filterNoTitle; // Set default state
        filterNoTitleCheckbox.addEventListener('change', this.handleFilterChange.bind(this));
      }
      
      // Load email history from localStorage
      this.loadEmailHistory();
    },
    
    // Load email history from localStorage
    loadEmailHistory: function() {
      try {
        const history = localStorage.getItem('orgChartEmailHistory');
        if (history) {
          const emails = JSON.parse(history);
          const datalist = document.getElementById('emailHistory');
          if (datalist) {
            datalist.innerHTML = '';
            // Add up to 10 most recent emails
            emails.slice(0, 10).forEach(email => {
              const option = document.createElement('option');
              option.value = email;
              datalist.appendChild(option);
            });
          }
        }
      } catch (error) {
        console.error('Error loading email history:', error);
      }
    },
    
    // Save email to history
    saveEmailToHistory: function(email) {
      try {
        let history = [];
        const stored = localStorage.getItem('orgChartEmailHistory');
        if (stored) {
          history = JSON.parse(stored);
        }
        
        // Remove email if it already exists (to move it to front)
        history = history.filter(e => e !== email);
        
        // Add email to beginning
        history.unshift(email);
        
        // Keep only last 10 emails
        history = history.slice(0, 10);
        
        // Save to localStorage
        localStorage.setItem('orgChartEmailHistory', JSON.stringify(history));
        
        // Reload the datalist
        this.loadEmailHistory();
      } catch (error) {
        console.error('Error saving email to history:', error);
      }
    },
    
    // Clear email history
    clearEmailHistory: function() {
      if (confirm('Clear all saved email addresses?')) {
        try {
          localStorage.removeItem('orgChartEmailHistory');
          const datalist = document.getElementById('emailHistory');
          if (datalist) {
            datalist.innerHTML = '';
          }
          // Also clear the current input
          const input = document.getElementById('startingEmail');
          if (input) {
            input.value = '';
          }
        } catch (error) {
          console.error('Error clearing email history:', error);
        }
      }
    },
    
    // Handle department filter change
    handleFilterChange: function(event) {
      // Update the appropriate filter based on which checkbox changed
      if (event.target.id === 'filterNoDept') {
        filterNoDepartment = event.target.checked;
      } else if (event.target.id === 'filterInterns') {
        filterInterns = event.target.checked;
      } else if (event.target.id === 'filterNoTitle') {
        filterNoTitle = event.target.checked;
      }
      
      // Reapply filter and update visualization if data exists
      if (hierarchyData && orgData.length > 0) {
        this.applyFilters();
        this.rebuildVisualization();
      }
    },
    
    // Apply filters to organization data
    applyFilters: function() {
      filteredOrgData = orgData.filter(user => {
        // Apply department filter
        if (filterNoDepartment && (!user.department || user.department === 'Unknown' || user.department.trim() === '')) {
          return false;
        }

        // Apply title filter
        if (filterNoTitle && (!user.title || user.title.trim() === '')) {
          return false;
        }

        // Apply intern filter (check if title contains "intern" - case insensitive)
        if (filterInterns && user.title && user.title.toLowerCase().includes('intern')) {
          return false;
        }
        
        return true;
      });
    },
    
    // Rebuild hierarchy with filtered data
    rebuildHierarchy: function(node) {
      if (!node) return null;
      
      // Check if this user should be included
      const userEmail = node.email;
      const userData = filteredOrgData.find(u => u.email === userEmail);
      
      if (!userData) {
        // User is filtered out, but check if they have children we need to keep
        if (node.children && node.children.length > 0) {
          const validChildren = [];
          for (const child of node.children) {
            const rebuiltChild = this.rebuildHierarchy(child);
            if (rebuiltChild) {
              validChildren.push(rebuiltChild);
            }
          }
          // If this filtered user has valid children, we need to keep them
          // but we'll promote the children up the hierarchy
          return validChildren.length > 0 ? validChildren : null;
        }
        return null;
      }
      
      // User is not filtered, rebuild their node
      const newNode = {
        id: node.id,
        name: node.name,
        email: node.email,
        title: node.title,
        department: node.department,
        location: node.location,
        hasLicense: node.hasLicense,
        aiEngagement: node.aiEngagement
      };
      
      // Process children
      if (node.children && node.children.length > 0) {
        const validChildren = [];
        for (const child of node.children) {
          const rebuiltChild = this.rebuildHierarchy(child);
          if (rebuiltChild) {
            // Handle case where child might be an array (promoted children)
            if (Array.isArray(rebuiltChild)) {
              validChildren.push(...rebuiltChild);
            } else {
              validChildren.push(rebuiltChild);
            }
          }
        }
        if (validChildren.length > 0) {
          newNode.children = validChildren;
        }
      }
      
      return newNode;
    },
    
    // Rebuild and refresh visualization
    rebuildVisualization: function() {
      const filteredHierarchy = this.rebuildHierarchy(hierarchyData);
      
      if (filteredHierarchy) {
        // Handle case where root might be filtered out
        let finalHierarchy = filteredHierarchy;
        if (Array.isArray(filteredHierarchy)) {
          // If root was filtered, create a virtual root
          finalHierarchy = {
            name: 'Organization',
            email: 'root@org',
            department: 'Organization',
            hasLicense: false,
            children: filteredHierarchy
          };
        }
        
        // Update statistics with filtered data
        this.showStatistics();
        
        // Update visualization
        OrgChart.createVisualization(finalHierarchy, filteredOrgData, licensedEmails);
      }
    },
    
    // Handle CSV file upload
    handleCsvUpload: function(event) {
      const file = event.target.files[0];
      if (!file) return;
      
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const csvText = e.target.result;
          const lines = csvText.split('\n');
          
          // Clear previous data
          licensedEmails.clear();
          licensedUsernames.clear();
          aiEngagementMap.clear();
          
          // Find column indices (case-insensitive)
          const headers = lines[0].toLowerCase().split(',').map(h => h.trim());
          const emailIndex = headers.indexOf('email');
          const engagementIndex = headers.indexOf('avg_daily_messages_total');
          
          if (emailIndex === -1) {
            throw new Error('No "email" column found in CSV');
          }
          
          // Process each row
          for (let i = 1; i < lines.length; i++) {
            const row = lines[i].split(',');
            if (row.length > emailIndex) {
              const email = row[emailIndex].trim().toLowerCase();
              if (email) {
                licensedEmails.add(email);
                const username = email.split('@')[0];
                if (username) {
                  licensedUsernames.add(username);

                  // If engagement data exists, store it keyed by username
                  if (engagementIndex !== -1 && row.length > engagementIndex) {
                    const avgMessages = parseFloat(row[engagementIndex].trim());
                    if (!isNaN(avgMessages)) {
                      // Divide by 3 to get engagement score
                      const engagementScore = avgMessages / 3;
                      aiEngagementMap.set(username, engagementScore);
                    }
                  }
                }
              }
            }
          }
          
          // Update UI
          document.getElementById('csvUpload').classList.add('has-file');
          const statusMessage = engagementIndex !== -1 ? 
            `✔ Loaded ${licensedEmails.size} licensed email addresses with AI engagement data` :
            `✔ Loaded ${licensedEmails.size} licensed email addresses`;
          document.getElementById('csvStatus').innerHTML = `
            <div class="status-success">
              ${statusMessage}
            </div>
          `;
          document.getElementById('step3Number').classList.add('step-complete');
          
        } catch (error) {
          document.getElementById('csvStatus').innerHTML = `
            <div class="status-error">
              ✗ Error parsing CSV: ${error.message}
            </div>
          `;
        }
      };
      
      reader.readAsText(file);
    },
    
    // Fetch organization data
    fetchOrgData: async function() {
      // Get input values
      accessToken = document.getElementById('graphToken').value.trim();
      startingEmail = document.getElementById('startingEmail').value.trim();
      maxUsers = parseInt(document.getElementById('maxUsers').value) || 100;
      const maxDepth = parseInt(document.getElementById('maxDepth').value) || 4;
      
      // Validate inputs
      if (!accessToken) {
        this.showError('fetchStatus', 'Please enter a Microsoft Graph access token');
        return;
      }
      
      if (!startingEmail) {
        this.showError('fetchStatus', 'Please enter a starting user email address');
        return;
      }

      // Force email to use company domain
      const username = startingEmail.split('@')[0].toLowerCase();
      startingEmail = `${username}@${orgDomain}`;
      const emailInput = document.getElementById('startingEmail');
      if (emailInput) {
        emailInput.value = startingEmail;
      }

      // Reset data
      orgData = [];
      hierarchyData = null;
      
      // Show progress
      document.getElementById('fetchButton').disabled = true;
      document.getElementById('fetchProgress').style.display = 'block';
      this.updateProgress(0, 'Fetching starting user...');
      
      try {
        // Mark step 1 as complete
        document.getElementById('step1Number').classList.add('step-complete');
        document.getElementById('step2Number').classList.add('step-complete');
        
        // Fetch starting user
        const startingUser = await this.fetchUser(startingEmail);
        if (!startingUser) {
          throw new Error('Starting user not found');
        }
        
        // Save successful email to history
        this.saveEmailToHistory(startingEmail);
        
        // Build organization hierarchy with depth limit
        hierarchyData = await this.buildHierarchy(startingUser, new Set(), 0, maxDepth);
        
        // Apply filters
        this.applyFilters();
        
        // Rebuild hierarchy with filtered data
        const filteredHierarchy = this.rebuildHierarchy(hierarchyData);
        
        let finalHierarchy = filteredHierarchy;
        if (!filteredHierarchy) {
          throw new Error('No data after filtering');
        }
        
        // Handle case where root might be filtered out
        if (Array.isArray(filteredHierarchy)) {
          finalHierarchy = {
            name: 'Organization',
            email: 'root@org',
            department: 'Organization',
            hasLicense: false,
            children: filteredHierarchy
          };
        }
        
        // Mark step 4 as complete
        document.getElementById('step4Number').classList.add('step-complete');
        
        // Show success message
        const depthLabel = maxDepth === 2 ? '1-2 levels' :
                          maxDepth === 3 ? '2-3 levels' :
                          maxDepth === 4 ? '3-5 levels' :
                          'unlimited depth';
        let filterInfo = '';
        if ((filterNoDepartment || filterInterns || filterNoTitle) && filteredOrgData.length < orgData.length) {
          filterInfo = ` - ${filteredOrgData.length} shown after filtering`;
        }
        document.getElementById('fetchStatus').innerHTML = `
          <div class="status-success">
            ✔ Successfully fetched ${orgData.length} users (${depthLabel})${filterInfo}
          </div>
        `;
        
        // Auto-collapse the load section
        const loadSection = document.getElementById('loadSection');
        const loadSectionIcon = document.getElementById('loadSectionIcon');
        if (loadSection && loadSectionIcon) {
          loadSection.style.display = 'none';
          loadSectionIcon.textContent = '▶';
        }
        
        // Show statistics
        this.showStatistics();
        
        // Show chart controls section
        document.getElementById('chartControlsSection').style.display = 'block';
        
        // Create visualization
        OrgChart.createVisualization(finalHierarchy, filteredOrgData, licensedEmails);
        
      } catch (error) {
        console.error('Error fetching org data:', error);
        this.showError('fetchStatus', `Error: ${error.message}`);
      } finally {
        document.getElementById('fetchButton').disabled = false;
        document.getElementById('fetchProgress').style.display = 'none';
      }
    },
    
    // Fetch a single user by email
    fetchUser: async function(email) {
      try {
        const response = await fetch(
          `${graphBaseUrl}/users/${encodeURIComponent(email)}?$select=id,displayName,mail,jobTitle,department,city,country,officeLocation`,
          {
            headers: {
              'Authorization': `Bearer ${accessToken}`,
              'Content-Type': 'application/json'
            }
          }
        );

        if (!response.ok) {
          if (response.status === 401) {
            throw new Error('Invalid or expired access token');
          }
          if (response.status === 404) {
            return null;
          }
          throw new Error(`Failed to fetch user: ${response.status}`);
        }

        const userData = await response.json();
        return this.processUserData(userData);

      } catch (error) {
        console.error(`Error fetching user ${email}:`, error);
        throw error;
      }
    },
    
    // Process user data
    processUserData: function(userData) {
      const email = (userData.mail || userData.userPrincipalName || '').toLowerCase();
      const username = email.split('@')[0];
      const user = {
        id: userData.id,
        name: userData.displayName || 'Unknown',
        email: email,
        title: userData.jobTitle || '',
        department: userData.department || 'Unknown',
        location: userData.city || userData.officeLocation || '',
        hasLicense: false,
        aiEngagement: 0
      };

      // Check if user has license using username
      if (username && licensedUsernames.has(username)) {
        user.hasLicense = true;
      }

      // Check if user has AI engagement data using username
      if (username && aiEngagementMap.has(username)) {
        user.aiEngagement = aiEngagementMap.get(username);
      }
      
      // Add to org data
      orgData.push(user);
      
      return user;
    },
    
    // Calculate star rating from engagement score
    getStarRating: function(engagementScore) {
      if (engagementScore === 0) {
        return { stars: '—', count: 0, label: 'No AI usage' };
      } else if (engagementScore <= 1) {
        return { stars: '★', count: 1, label: 'Light user' };
      } else if (engagementScore <= 2) {
        return { stars: '★★', count: 2, label: 'Regular user' };
      } else if (engagementScore <= 3) {
        return { stars: '★★★', count: 3, label: 'Active user' };
      } else if (engagementScore <= 4) {
        return { stars: '★★★★', count: 4, label: 'Heavy user' };
      } else if (engagementScore <= 5) {
        return { stars: '★★★★★', count: 5, label: 'Power user' };
      } else {
        return { stars: '★★★★★🚀', count: 5, label: 'Super power user' };
      }
    },
    
    // Build organization hierarchy
    buildHierarchy: async function(user, visitedIds, depth = 0, maxDepth = 999) {
      // Check if we've reached max users or max depth
      if (orgData.length >= maxUsers || depth >= maxDepth) {
        return null;
      }
      
      // Check if we've already visited this user
      if (visitedIds.has(user.id)) {
        return null;
      }
      
      visitedIds.add(user.id);
      
      // Create node for this user
      const node = {
        id: user.id,
        name: user.name,
        email: user.email,
        title: user.title,
        department: user.department,
        location: user.location,
        hasLicense: user.hasLicense,
        aiEngagement: user.aiEngagement,
        children: []
      };
      
      // Update progress
      const depthText = maxDepth === 999 ? '' : ` (Level ${depth + 1}/${maxDepth})`;
      this.updateProgress(
        (orgData.length / maxUsers) * 100,
        `Fetching organization structure... (${orgData.length}/${maxUsers} users${depthText})`
      );
      
      // Fetch direct reports only if we haven't reached max depth
      if (depth < maxDepth - 1) {
        try {
          const response = await fetch(
            `${graphBaseUrl}/users/${user.id}/directReports?$select=id,displayName,mail,jobTitle,department,city,country,officeLocation`,
            {
              headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
              }
            }
          );

          if (response.status === 401) {
            throw new Error('Invalid or expired access token');
          }

          if (response.ok) {
            const data = await response.json();
            const reports = data.value || [];

            // Process each direct report
            for (const reportData of reports) {
              if (orgData.length >= maxUsers) break;

              const report = this.processUserData(reportData);
              if (report) {
                const childNode = await this.buildHierarchy(report, visitedIds, depth + 1, maxDepth);
                if (childNode) {
                  node.children.push(childNode);
                }
              }
            }
          } else {
            console.error(`Failed to fetch direct reports for ${user.name}: ${response.status}`);
          }
        } catch (error) {
          console.error(`Error fetching direct reports for ${user.name}:`, error);
          throw error;
        }
      }
      
      // If no children, delete the empty array to help D3 rendering
      if (node.children.length === 0) {
        delete node.children;
      }
      
      return node;
    },
    
    // Update progress bar
    updateProgress: function(percent, text) {
      document.getElementById('progressFill').style.width = `${percent}%`;
      document.getElementById('progressText').textContent = text;
    },
    
    // Show error message
    showError: function(elementId, message) {
      document.getElementById(elementId).innerHTML = `
        <div class="status-error">
          ✗ ${message}
        </div>
      `;
    },
    
    // Show statistics
    showStatistics: function() {
      // Use filtered data for statistics
      const dataToUse = filteredOrgData.length > 0 ? filteredOrgData : orgData;
      
      const licensed = dataToUse.filter(u => u.hasLicense).length;
      const unlicensed = dataToUse.length - licensed;
      
      // Calculate AI engagement statistics
      let engagementStats = { none: 0, light: 0, regular: 0, active: 0, heavy: 0, power: 0, super: 0 };
      let totalEngaged = 0;
      dataToUse.forEach(user => {
        if (user.aiEngagement !== undefined) {
          const rating = this.getStarRating(user.aiEngagement);
          if (rating.count === 0) engagementStats.none++;
          else if (rating.count === 1) engagementStats.light++;
          else if (rating.count === 2) engagementStats.regular++;
          else if (rating.count === 3) engagementStats.active++;
          else if (rating.count === 4) engagementStats.heavy++;
          else if (rating.count === 5 && user.aiEngagement <= 5) engagementStats.power++;
          else if (user.aiEngagement > 5) engagementStats.super++;
          
          if (user.aiEngagement > 0) totalEngaged++;
        }
      });
      
      // Calculate department statistics
      const deptStats = {};
      dataToUse.forEach(user => {
        const dept = user.department || 'Unknown';
        if (!deptStats[dept]) {
          deptStats[dept] = { total: 0, licensed: 0, engaged: 0, totalEngagement: 0 };
        }
        deptStats[dept].total++;
        if (user.hasLicense) {
          deptStats[dept].licensed++;
        }
        if (user.aiEngagement > 0) {
          deptStats[dept].engaged++;
          deptStats[dept].totalEngagement += user.aiEngagement;
        }
      });
      
      // Sort departments by total users
      const sortedDepts = Object.entries(deptStats)
        .sort((a, b) => b[1].total - a[1].total)
        .map(([dept, stats]) => ({
          name: dept,
          total: stats.total,
          licensed: stats.licensed,
          engaged: stats.engaged,
          avgEngagement: stats.engaged > 0 ? (stats.totalEngagement / stats.engaged).toFixed(1) : 0,
          percent: Math.round((stats.licensed / stats.total) * 100)
        }));
      
      // Create department rows
      const deptRows = sortedDepts.slice(0, 10).map(dept => {
        const barColor = dept.percent >= 75 ? 'var(--success)' : 
                        dept.percent >= 50 ? 'var(--warning)' : 
                        'var(--danger)';
        // Escape the department name for safe use in HTML attribute
        const escapedDeptName = dept.name.replace(/'/g, '&#39;').replace(/"/g, '&quot;');
        
        // Calculate star display for average engagement
        const avgStars = dept.avgEngagement > 0 ? this.getStarRating(parseFloat(dept.avgEngagement)).stars : '—';
        
        return `
          <tr class="dept-row" data-department="${escapedDeptName}" onclick="OrgChart.highlightDepartment(this.dataset.department)">
            <td style="font-weight: 500;">${dept.name}</td>
            <td style="text-align: center;">${dept.licensed}/${dept.total}</td>
            <td style="width: 100px;">
              <div class="mini-progress">
                <div class="mini-progress-fill" style="width: ${dept.percent}%; background: ${barColor};"></div>
                <span class="mini-progress-text">${dept.percent}%</span>
              </div>
            </td>
            ${totalEngaged > 0 ? `<td style="text-align: center;">${dept.engaged > 0 ? `${dept.engaged} (${avgStars})` : '—'}</td>` : ''}
          </tr>
        `;
      }).join('');
      
      // Add filter status message if filtering
      let filterMessage = '';
      if ((filterNoDepartment || filterInterns || filterNoTitle) && orgData.length > dataToUse.length) {
        const filtered = orgData.length - dataToUse.length;
        const filterReasons = [];
        if (filterNoDepartment) filterReasons.push('no department');
        if (filterInterns) filterReasons.push('interns');
        if (filterNoTitle) filterReasons.push('no title');
        
        filterMessage = `<div class="filter-status">
          <span style="color: var(--info);">ℹ️ Showing ${dataToUse.length} of ${orgData.length} users (${filtered} filtered: ${filterReasons.join(', ')})</span>
        </div>`;
      }
      
      const statsHtml = `
        ${filterMessage}
        <div class="stats-grid">
          <div class="stat-card">
            <div class="stat-title">Overall Adoption</div>
            <div class="stat-number">${Math.round(licensed/dataToUse.length*100)}%</div>
            <div class="stat-subtitle">${licensed} of ${dataToUse.length} users</div>
            <div class="progress-bar" style="margin-top: 0.5rem;">
              <div class="progress-fill" style="width: ${(licensed/dataToUse.length)*100}%; background: var(--success);"></div>
            </div>
          </div>
          
          <div class="stat-card">
            <div class="stat-row">
              <span>Total Users:</span>
              <strong>${dataToUse.length}</strong>
            </div>
            <div class="stat-row">
              <span>Licensed:</span>
              <strong style="color: var(--success);">${licensed}</strong>
            </div>
            <div class="stat-row">
              <span>Unlicensed:</span>
              <strong style="color: var(--danger);">${unlicensed}</strong>
            </div>
            <div class="stat-row">
              <span>AI Active:</span>
              <strong style="color: var(--info);">${totalEngaged}</strong>
            </div>
          </div>
        </div>
        
        ${totalEngaged > 0 ? `
        <div class="engagement-stats">
          <h4 style="color: var(--primary); margin-bottom: 0.75rem;">AI Engagement Levels</h4>
          <div style="display: flex; flex-wrap: wrap; gap: 0.5rem; font-size: 0.85rem;">
            ${engagementStats.light > 0 ? `<span style="padding: 0.25rem 0.5rem; background: var(--bg); border: 1px solid var(--border); border-radius: var(--radius);">★ Light: ${engagementStats.light}</span>` : ''}
            ${engagementStats.regular > 0 ? `<span style="padding: 0.25rem 0.5rem; background: var(--bg); border: 1px solid var(--border); border-radius: var(--radius);">★★ Regular: ${engagementStats.regular}</span>` : ''}
            ${engagementStats.active > 0 ? `<span style="padding: 0.25rem 0.5rem; background: var(--bg); border: 1px solid var(--border); border-radius: var(--radius);">★★★ Active: ${engagementStats.active}</span>` : ''}
            ${engagementStats.heavy > 0 ? `<span style="padding: 0.25rem 0.5rem; background: var(--bg); border: 1px solid var(--border); border-radius: var(--radius);">★★★★ Heavy: ${engagementStats.heavy}</span>` : ''}
            ${engagementStats.power > 0 ? `<span style="padding: 0.25rem 0.5rem; background: var(--bg); border: 1px solid var(--border); border-radius: var(--radius);">★★★★★ Power: ${engagementStats.power}</span>` : ''}
            ${engagementStats.super > 0 ? `<span style="padding: 0.25rem 0.5rem; background: var(--bg); border: 1px solid var(--border); border-radius: var(--radius);">★★★★★🚀 Super: ${engagementStats.super}</span>` : ''}
            ${engagementStats.none > 0 ? `<span style="padding: 0.25rem 0.5rem; background: rgba(0,0,0,0.05); border: 1px solid var(--border); border-radius: var(--radius); color: var(--text-light);">— No usage: ${engagementStats.none}</span>` : ''}
          </div>
        </div>
        ` : ''}
        
        <div class="dept-stats-table">
          <h4>License Adoption by Department</h4>
          <div class="dept-hint">Click a department to highlight in chart</div>
          <table>
            <thead>
              <tr>
                <th>Department</th>
                <th>Licensed</th>
                <th>Adoption</th>
                ${totalEngaged > 0 ? '<th>AI Active</th>' : ''}
              </tr>
            </thead>
            <tbody>
              ${deptRows}
            </tbody>
          </table>
          ${sortedDepts.length > 10 ? `<div class="table-footer">...and ${sortedDepts.length - 10} more departments</div>` : ''}
        </div>
      `;
      
      document.getElementById('statsSection').style.display = 'block';
      document.getElementById('statsContent').innerHTML = statsHtml;
    }
  }; 
})();

// Expose GraphAPI to global scope for inline event handlers
window.GraphAPI = GraphAPI;