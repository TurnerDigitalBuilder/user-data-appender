// D3.js Organization Chart Visualization
const OrgChart = (function() {
  'use strict';
  
  // Private variables
  let svg = null;
  let g = null;
  let root = null;
  let treemap = null;
  let zoom = null;
  let i = 0;
  let orgData = [];
  let licensedEmails = new Set();
  let highlightedDepartment = null;
  let highlightedNodes = null;
  let licenseHighlightActive = false;
  let currentExpandLevel = 1;

  // Configuration
  const config = {
    horizontalSpacing: 200,
    verticalSpacing: 1,
    nodeRadius: 10,
    duration: 750,
    margin: { top: 20, right: 120, bottom: 20, left: 120 },
    showEngagementStars: true,
    dimMode: 'dim'
  };
  
  // Color scales
  let departmentColorScale = null;
  let levelColorScale = null;

  // Helper functions
  function getAllChildren(node) {
    return [
      ...(node.children || []),
      ...(node._children || [])
    ];
  }

  function countDescendants(node) {
    return getAllChildren(node)
      .reduce((sum, child) => sum + 1 + countDescendants(child), 0);
  }

  function countLicensedDescendants(node) {
    return getAllChildren(node)
      .reduce((sum, child) => {
        const isLicensed = child.data && child.data.hasLicense ? 1 : 0;
        return sum + isLicensed + countLicensedDescendants(child);
      }, 0);
  }
  
  // Public API
  return {
    // Initialize the visualization module
    initialize: function() {
      // Set up event listeners
      const horizontalSlider = document.getElementById('horizontalSpacing');
      if (horizontalSlider) {
        horizontalSlider.addEventListener('input', (e) => {
          config.horizontalSpacing = parseInt(e.target.value);
          document.getElementById('horizontalValue').textContent = e.target.value;
        });
      }
      
      const verticalSlider = document.getElementById('verticalSpacing');
      if (verticalSlider) {
        verticalSlider.addEventListener('input', (e) => {
          config.verticalSpacing = parseFloat(e.target.value);
          document.getElementById('verticalValue').textContent = e.target.value;
        });
      }

      const showStarsCheckbox = document.getElementById('showStars');
      if (showStarsCheckbox) {
        config.showEngagementStars = showStarsCheckbox.checked;
        showStarsCheckbox.addEventListener('change', (e) => {
          config.showEngagementStars = e.target.checked;
          if (root) this.update(root);
        });
      }

      const dimModeSelect = document.getElementById('dimMode');
      if (dimModeSelect) {
        config.dimMode = dimModeSelect.value;
        dimModeSelect.addEventListener('change', (e) => {
          config.dimMode = e.target.value;
          if (root) this.update(root);
        });
      }
    },
    
    // Create the visualization
    createVisualization: function(hierarchyData, userData, licenses) {
      if (!hierarchyData) {
        console.error('No hierarchy data provided');
        return;
      }
      
      // Reset any existing search or department highlights
      highlightedDepartment = null;
      highlightedNodes = null;
      licenseHighlightActive = false;
      document.querySelectorAll('.dept-row').forEach(row => row.classList.remove('active'));
      const searchInput = document.getElementById('searchInput');
      if (searchInput) searchInput.value = '';
      const clearBtn = document.getElementById('clearHighlightBtn');
      if (clearBtn) clearBtn.style.display = 'none';
      const licenseBtn = document.getElementById('licenseToggleBtn');
      if (licenseBtn) licenseBtn.classList.remove('active');

      orgData = userData;
      licensedEmails = licenses;

      // Set up color scales
      this.setupColorScales();
      
      // Get container dimensions
      const container = document.getElementById('treeContainer');
      const width = container.clientWidth - config.margin.left - config.margin.right;
      const height = container.clientHeight - config.margin.top - config.margin.bottom;
      
      // Clear previous visualization
      d3.select('#treeContainer').selectAll('*').remove();
      
      // Create SVG
      svg = d3.select('#treeContainer')
        .append('svg')
        .attr('width', container.clientWidth)
        .attr('height', container.clientHeight);
      
      // Create group for zoom
      g = svg.append('g')
        .attr('transform', `translate(${config.margin.left},${height / 2})`);
      
      // Set up zoom behavior
      zoom = d3.zoom()
        .scaleExtent([0.1, 3])
        .on('zoom', (event) => {
          g.attr('transform', event.transform);
        });
      
      svg.call(zoom);

      // Hide selected user panel when clicking on empty space
      svg.on('click', () => {
        const section = document.getElementById('selectedUserSection');
        if (section) {
          section.style.display = 'none';
          section.innerHTML = '';
        }
      });
      
      // Create tree layout
      treemap = d3.tree().size([height, width]);
      
      // Create hierarchy
      root = d3.hierarchy(hierarchyData);
      root.x0 = height / 2;
      root.y0 = 0;
      
      // Store original positions
      root.descendants().forEach(d => {
        d.x0 = d.x;
        d.y0 = d.y;
      });
      
      // Collapse after the second level initially
      if (root.children) {
        root.children.forEach(this.collapse);
      }
      currentExpandLevel = 1;

      // Initial render
      this.update(root);

      // Show controls and panels
      document.getElementById('legend').style.display = 'block';
      const selectedSection = document.getElementById('selectedUserSection');
      if (selectedSection) {
        selectedSection.style.display = 'none';
        selectedSection.innerHTML = '';
      }
      document.getElementById('statsPanel').style.display = 'block';
      document.getElementById('bottomControls').style.display = 'flex';

      this.updateLegend();
      this.updateStatsPanel();
    },
    
    // Setup color scales
    setupColorScales: function() {
      const departments = [...new Set(orgData.map(u => u.department))];
      // Use a color scheme that avoids reds to prevent confusion with license borders
      const colorPalette = [
        '#1f77b4', // blue
        '#2ca02c', // green
        '#ff7f0e', // orange
        '#9467bd', // purple
        '#8c564b', // brown
        '#e377c2', // pink
        '#7f7f7f', // gray
        '#bcbd22', // olive
        '#17becf', // cyan
        '#aec7e8', // light blue
        '#ffbb78', // light orange
        '#98df8a', // light green
        '#c5b0d5', // light purple
        '#c49c94', // light brown
        '#f7b6d2', // light pink
        '#c7c7c7', // light gray
        '#dbdb8d', // light olive
        '#9edae5'  // light cyan
      ];
      
      departmentColorScale = d3.scaleOrdinal()
        .domain(departments)
        .range(colorPalette);
      
      levelColorScale = d3.scaleSequential()
        .domain([0, 5])
        .interpolator(d3.interpolateBlues);
    },
    
    // Update the tree
    update: function(source) {
      if (!root || !treemap) return;
      
      // Compute the new tree layout
      const treeData = treemap(root);
      const nodes = treeData.descendants();
      const links = treeData.links();
      
      // Normalize for fixed-depth and apply spacing
      nodes.forEach(d => {
        d.y = d.depth * config.horizontalSpacing;
        d.x = d.x * config.verticalSpacing;
      });
      
      // -------------------- NODES --------------------
      const node = g.selectAll('g.node')
        .data(nodes, d => d.id || (d.id = ++i));
      
      // Enter new nodes at the parent's previous position
      const nodeEnter = node.enter().append('g')
        .attr('class', 'node')
        .attr('transform', d => `translate(${source.y0 || 0},${source.x0 || 0})`)
        .style('display', d => this.getNodeDisplay(d))
        .on('click', (event, d) => {
          // Prevent background click handler from firing when selecting a node
          event.stopPropagation();
          this.click(event, d);
        })
        .on('mouseover', (event, d) => this.showTooltip(event, d))
        .on('mouseout', () => this.hideTooltip());
      
      // Add circles for nodes
      nodeEnter.append('circle')
        .attr('r', 1e-6)
        .style('fill', d => this.getNodeColor(d))
        .style('stroke', d => this.getNodeStroke(d))
        .style('opacity', d => this.getNodeOpacity(d));
      
      // Add labels for nodes - include star rating if available
      nodeEnter.append('text')
        .attr('dy', '.35em')
        .attr('x', d => d.children || d._children ? -13 : 13)
        .attr('text-anchor', d => d.children || d._children ? 'end' : 'start')
        .text(d => this.getNodeLabel(d))
        .style('fill-opacity', 1e-6);
      
      // Add title (job title) as second line
      nodeEnter.append('text')
        .attr('class', 'title')
        .attr('dy', '1.5em')
        .attr('x', d => d.children || d._children ? -13 : 13)
        .attr('text-anchor', d => d.children || d._children ? 'end' : 'start')
        .text(d => d.data.title || '')
        .style('fill-opacity', 1e-6);
      
      // UPDATE
      const nodeUpdate = nodeEnter.merge(node);

      // Transition to the proper position for the node
      nodeUpdate.transition()
        .duration(config.duration)
        .attr('transform', d => `translate(${d.y},${d.x})`);

      // Disable interactions for dimmed nodes
      nodeUpdate
        .style('display', d => this.getNodeDisplay(d))
        .style('pointer-events', d =>
          !highlightedNodes || highlightedNodes.has(d.data.email) ? 'all' : 'none')
        .style('cursor', d =>
          !highlightedNodes || highlightedNodes.has(d.data.email) ? 'pointer' : 'default');
      
      // Update the node attributes and style
      nodeUpdate.select('circle')
        .attr('r', config.nodeRadius)
        .style('fill', d => this.getNodeColor(d))
        .style('stroke', d => this.getNodeStroke(d))
        .style('opacity', d => this.getNodeOpacity(d));
      
      nodeUpdate.select('text')
        .text(d => this.getNodeLabel(d))
        .style('fill-opacity', d => this.getNodeOpacity(d));
      
      nodeUpdate.select('text.title')
        .style('fill-opacity', d => this.getNodeOpacity(d) * 0.7);
      
      // EXIT
      const nodeExit = node.exit().transition()
        .duration(config.duration)
        .attr('transform', d => `translate(${source.y},${source.x})`)
        .remove();
      
      nodeExit.select('circle')
        .attr('r', 1e-6);
      
      nodeExit.select('text')
        .style('fill-opacity', 1e-6);
      
      // -------------------- LINKS --------------------
      const link = g.selectAll('path.link')
        .data(links, d => d.target.id);
      
      // Enter new links at the parent's previous position
      const linkEnter = link.enter().insert('path', 'g')
        .attr('class', 'link')
        .style('display', d => this.getLinkDisplay(d))
        .style('opacity', d => this.getLinkOpacity(d))
        .attr('d', d => {
          const o = { x: source.x0 || 0, y: source.y0 || 0 };
          return this.diagonal(o, o);
        });

      // UPDATE
      const linkUpdate = linkEnter.merge(link);

      // Transition back to the parent element position
      linkUpdate
        .style('display', d => this.getLinkDisplay(d))
        .transition()
        .duration(config.duration)
        .attr('d', d => this.diagonal(d.source, d.target))
        .style('opacity', d => this.getLinkOpacity(d));
      
      // EXIT
      link.exit().transition()
        .duration(config.duration)
        .attr('d', d => {
          const o = { x: source.x, y: source.y };
          return this.diagonal(o, o);
        })
        .remove();
      
      // Store the old positions for transition
      nodes.forEach(d => {
        d.x0 = d.x;
        d.y0 = d.y;
      });

      // Refresh statistics to reflect current view
      this.updateStatsPanel();
    },
    
    // Creates a curved (diagonal) path from parent to child nodes
    diagonal: function(source, target) {
      // Handle both cases: when called with link object or separate source/target
      const s = source.source ? source.source : source;
      const d = source.target ? source.target : target;
      
      // Safety check for valid coordinates
      if (!s || !d || 
          s.y === undefined || s.x === undefined || 
          d.y === undefined || d.x === undefined) {
        return '';
      }
      
      return `M ${s.y} ${s.x}
              C ${(s.y + d.y) / 2} ${s.x},
                ${(s.y + d.y) / 2} ${d.x},
                ${d.y} ${d.x}`;
    },
    
    // Toggle children on click
    click: function(event, d) {
      if (d.children) {
        d._children = d.children;
        d.children = null;
      } else {
        d.children = d._children;
        d._children = null;
      }
      this.update(d);
      this.updateSelectedUser(d);
    },
    
    // Collapse node
    collapse: function(d) {
      if (d.children) {
        d._children = d.children;
        d._children.forEach(OrgChart.collapse);
        d.children = null;
      }
    },

    // Expand next tier of nodes
    expandOneLevel: function() {
      if (!root) return;

      let expanded = false;

      function traverse(node, depth) {
        if (depth <= currentExpandLevel && node._children) {
          node.children = node._children;
          node._children = null;
          expanded = true;
        }

        const children = node.children ? node.children.slice() : [];
        const hidden = node._children ? node._children.slice() : [];
        [...children, ...hidden].forEach(child => traverse(child, depth + 1));
      }

      traverse(root, 0);

      if (expanded) {
        currentExpandLevel++;
        this.update(root);
      }
    },

    // Collapse the deepest expanded tier
    collapseOneLevel: function() {
      if (!root) return;

      const nodes = root.descendants();
      let deepestWithChildren = -1;

      nodes.forEach((node) => {
        if (node.children && node.children.length > 0) {
          deepestWithChildren = Math.max(deepestWithChildren, node.depth);
        }
      });

      if (deepestWithChildren < 0) return;

      let collapsed = false;
      nodes.forEach((node) => {
        if (node.depth === deepestWithChildren && node.children && node.children.length > 0) {
          this.collapse(node);
          collapsed = true;
        }
      });

      if (!collapsed) return;

      let newDeepest = -1;
      root.descendants().forEach((node) => {
        if (node.children && node.children.length > 0) {
          newDeepest = Math.max(newDeepest, node.depth);
        }
      });

      currentExpandLevel = Math.max(0, newDeepest + 1);
      this.update(root);
    },

    // Expand all nodes
    expandAll: function() {
      if (!root) return;

      function expand(d) {
        if (d._children) {
          d.children = d._children;
          d._children = null;
        }
        if (d.children) {
          d.children.forEach(expand);
        }
      }

      expand(root);
      currentExpandLevel = root.height + 1;
      this.update(root);
    },

    // Collapse all nodes
    collapseAll: function() {
      if (!root) return;

      if (root.children) {
        root.children.forEach(this.collapse);
      }
      currentExpandLevel = 1;
      this.update(root);
    },
    
    // Reset view to center
    resetView: function() {
      if (!svg || !root) return;

      const nodes = root.descendants();
      let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
      nodes.forEach(d => {
        minX = Math.min(minX, d.x);
        maxX = Math.max(maxX, d.x);
        minY = Math.min(minY, d.y);
        maxY = Math.max(maxY, d.y);
      });

      const container = document.getElementById('treeContainer');
      const width = container.clientWidth - config.margin.left - config.margin.right;
      const height = container.clientHeight - config.margin.top - config.margin.bottom;

      const dx = maxX - minX;
      const dy = maxY - minY;
      const scale = Math.min(width / (dy || width), height / (dx || height));
      const tx = -minY * scale + (width - dy * scale) / 2 + config.margin.left;
      const ty = -minX * scale + (height - dx * scale) / 2 + config.margin.top;

      svg.transition()
        .duration(750)
        .call(zoom.transform, d3.zoomIdentity.translate(tx, ty).scale(scale));
    },

    // Download current graph data as CSV
    downloadCSV: function() {
      if (!orgData || orgData.length === 0) return;

      // Determine which users to export
      let dataToExport = orgData;
      if (highlightedNodes && highlightedNodes.size > 0) {
        dataToExport = orgData.filter(u => highlightedNodes.has(u.email));
      }

      const header = ['Name', 'Title', 'Department', 'Location', 'Email', 'License', 'AI Usage'];
      const rows = dataToExport.map(u => {
        const rating = u.aiEngagement !== undefined ? GraphAPI.getStarRating(u.aiEngagement) : null;
        const daily = u.aiEngagement !== undefined ? Math.round(u.aiEngagement * 3) : '';
        const aiUsage = rating ? `${rating.label} (~${daily}/day)` : '';
        return [
          u.name || '',
          u.title || '',
          u.department || '',
          u.location || '',
          u.email || '',
          u.hasLicense ? 'Yes' : 'No',
          aiUsage
        ];
      });

      const csv = [header, ...rows]
        .map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(','))
        .join('\n');

      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'org-data.csv';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    },

    // Update spacing
    updateSpacing: function(type, value) {
      if (type === 'horizontal') {
        config.horizontalSpacing = parseInt(value);
        document.getElementById('horizontalValue').textContent = value;
      } else if (type === 'vertical') {
        config.verticalSpacing = parseFloat(value);
        document.getElementById('verticalValue').textContent = value;
      }
      
      if (root) {
        this.update(root);
      }
    },
    
    // Update visualization with new color scheme
    updateVisualization: function() {
      if (root) {
        this.update(root);
        this.updateLegend();
      }
    },
    
    // Highlight department nodes
    highlightDepartment: function(department) {
      // If clicking the same department, clear the highlight
      if (highlightedDepartment === department) {
        this.clearHighlight();
      } else {
        highlightedDepartment = department;
        highlightedNodes = new Set(orgData
          .filter(u => u.department === department)
          .map(u => u.email));
        highlightedNodes = highlightedNodes.size > 0 ? highlightedNodes : null;
        const input = document.getElementById('searchInput');
        if (input) input.value = '';
        const searchClearBtn = document.getElementById('clearSearchBtn');
        if (searchClearBtn) searchClearBtn.style.display = 'none';
        licenseHighlightActive = false;
        const licenseBtn = document.getElementById('licenseToggleBtn');
        if (licenseBtn) licenseBtn.classList.remove('active');

        // Update active class on rows
        document.querySelectorAll('.dept-row').forEach(row => {
          if (row.dataset.department === department) {
            row.classList.add('active');
          } else {
            row.classList.remove('active');
          }
        });

        // Show clear highlight button
        const clearBtn = document.getElementById('clearHighlightBtn');
        if (clearBtn) {
          clearBtn.style.display = 'flex';
        }
      }

      // Update the visualization
      if (root) {
        this.update(root);
      }
    },

    // Toggle highlight of licensed users
    toggleLicenseHighlight: function() {
      if (licenseHighlightActive) {
        this.clearHighlight();
      } else {
        highlightedDepartment = null;
        highlightedNodes = new Set(orgData.filter(u => u.hasLicense).map(u => u.email));
        highlightedNodes = highlightedNodes.size > 0 ? highlightedNodes : null;
        licenseHighlightActive = true;
        document.querySelectorAll('.dept-row').forEach(row => row.classList.remove('active'));
        const input = document.getElementById('searchInput');
        if (input) input.value = '';
        const clearBtn = document.getElementById('clearHighlightBtn');
        if (clearBtn) clearBtn.style.display = 'flex';
        const searchClearBtn = document.getElementById('clearSearchBtn');
        if (searchClearBtn) searchClearBtn.style.display = 'none';
        const licenseBtn = document.getElementById('licenseToggleBtn');
        if (licenseBtn) licenseBtn.classList.add('active');
        if (root) this.update(root);
      }
    },

    // Search by name or title
    searchByNameTitle: function() {
      const input = document.getElementById('searchInput');
      if (!input) return;
      const raw = input.value.trim().toLowerCase();
      if (!raw) {
        this.clearHighlight();
        return;
      }

      // Split on quoted phrases or individual words
      const terms = raw.match(/"[^"]+"|\S+/g) || [];

      highlightedDepartment = null;
      highlightedNodes = new Set(orgData.filter(u => {
        const name = (u.name || '').toLowerCase();
        const title = (u.title || '').toLowerCase();
        return terms.some(t => {
          const token = t.replace(/^"|"$/g, '');
          return name.includes(token) || title.includes(token);
        });
      }).map(u => u.email));

      highlightedNodes = highlightedNodes.size > 0 ? highlightedNodes : null;
      licenseHighlightActive = false;
      const licenseBtn = document.getElementById('licenseToggleBtn');
      if (licenseBtn) licenseBtn.classList.remove('active');

      // Remove active class from department rows
      document.querySelectorAll('.dept-row').forEach(row => row.classList.remove('active'));

      // Show clear search button if we found matches
      const clearHighlightBtn = document.getElementById('clearHighlightBtn');
      if (clearHighlightBtn) {
        clearHighlightBtn.style.display = 'none';
      }
      const clearSearchBtn = document.getElementById('clearSearchBtn');
      if (clearSearchBtn) {
        clearSearchBtn.style.display = highlightedNodes && highlightedNodes.size > 0 ? 'flex' : 'none';
      }

      if (root) {
        this.update(root);
      }
    },

    // Clear highlight (department or search)
    clearHighlight: function() {
      highlightedDepartment = null;
      highlightedNodes = null;
      licenseHighlightActive = false;

      // Remove active class from all rows
      document.querySelectorAll('.dept-row').forEach(row => {
        row.classList.remove('active');
      });

      // Clear search input
      const input = document.getElementById('searchInput');
      if (input) input.value = '';

      // Hide clear buttons
      const clearHighlightBtn = document.getElementById('clearHighlightBtn');
      if (clearHighlightBtn) {
        clearHighlightBtn.style.display = 'none';
      }
      const clearSearchBtn = document.getElementById('clearSearchBtn');
      if (clearSearchBtn) {
        clearSearchBtn.style.display = 'none';
      }
      const licenseBtn = document.getElementById('licenseToggleBtn');
      if (licenseBtn) licenseBtn.classList.remove('active');

      // Update the visualization
      if (root) {
        this.update(root);
      }
    },

    // Get node opacity based on highlight
    getNodeOpacity: function(d) {
      if (!highlightedNodes) {
        return 1;
      }
      if (highlightedNodes.has(d.data.email)) {
        return 1;
      }

      switch (config.dimMode) {
        case 'transparent':
        case 'hidden':
          return 0;
        case 'dim':
        default:
          return 0.2;
      }
    },

    // Get link opacity based on highlight
    getLinkOpacity: function(d) {
      if (!highlightedNodes) {
        return 1;
      }
      // Show link if either source or target is in highlighted set
      const isHighlighted = (highlightedNodes.has(d.source.data.email) ||
              highlightedNodes.has(d.target.data.email));

      if (isHighlighted) {
        return 0.6;
      }

      switch (config.dimMode) {
        case 'transparent':
        case 'hidden':
          return 0;
        case 'dim':
        default:
          return 0.1;
      }
    },

    // Get node display based on highlight and dim mode
    getNodeDisplay: function(d) {
      if (!highlightedNodes || config.dimMode !== 'hidden') {
        return null;
      }

      return highlightedNodes.has(d.data.email) ? null : 'none';
    },

    // Get link display based on highlight and dim mode
    getLinkDisplay: function(d) {
      if (!highlightedNodes || config.dimMode !== 'hidden') {
        return null;
      }

      const isHighlighted = highlightedNodes.has(d.source.data.email) ||
        highlightedNodes.has(d.target.data.email);

      return isHighlighted ? null : 'none';
    },

    // Get node label with optional engagement stars
    getNodeLabel: function(d) {
      const name = d.data.name || 'Unknown';
      if (config.showEngagementStars && d.data.aiEngagement !== undefined && d.data.aiEngagement > 0) {
        const rating = GraphAPI.getStarRating(d.data.aiEngagement);
        return `${name} ${rating.stars}`;
      }
      return name;
    },

    // Get node color based on current color scheme
    getNodeColor: function(d) {
      const colorBy = document.getElementById('colorBy').value;
      
      if (colorBy === 'license') {
        return d.data.hasLicense ? '#00a854' : '#f0f0f0';
      } else if (colorBy === 'level') {
        return levelColorScale(Math.min(d.depth, 5));
      } else { // department
        return departmentColorScale(d.data.department || 'Unknown');
      }
    },
    
    // Get node stroke color
    getNodeStroke: function(d) {
      // Red border for no license, green for has license
      if (d.data.hasLicense) {
        return '#00a854'; // Green for licensed users
      }
      return '#dc3545'; // Red for unlicensed users
    },
    
    // Show tooltip
    showTooltip: function(event, d) {
      // Don't show tooltip for dimmed nodes when highlighting is active
      if (highlightedNodes && !highlightedNodes.has(d.data.email)) {
        return;
      }

      const tooltip = document.getElementById('tooltip');
      const licenseText = d.data.hasLicense ?
        '<span style="color: #00a854; font-weight: bold;">✔ Has ChatGPT License</span>' :
        '<span style="color: #dc3545; font-weight: bold;">✗ No ChatGPT License</span>';
      
      // Calculate AI engagement display
      let engagementText = '';
      if (d.data.aiEngagement !== undefined) {
          const rating = GraphAPI.getStarRating(d.data.aiEngagement);
          const dailyEngagements = Math.round(d.data.aiEngagement * 3);
          engagementText = `
            <div style="margin-top: 0.5rem; padding-top: 0.5rem; border-top: 1px solid #ddd;">
              <strong>AI Engagement:</strong> ${rating.stars}<br>
              <span style="color: var(--text-light); font-size: 0.85rem;">
                ${rating.label} (~${dailyEngagements} engagements/day)
              </span>
            </div>
          `;
        }

      const totalReports = countDescendants(d);
      const licensedReports = countLicensedDescendants(d);
      const licensePercent = totalReports > 0
        ? Number(((licensedReports / totalReports) * 100).toFixed(1))
        : 0;

      tooltip.innerHTML = `
        <strong>${d.data.name || 'Unknown'}</strong><br>
        ${d.data.title ? d.data.title + '<br>' : ''}
        ${d.data.department || 'No department'}<br>
        ${d.data.location ? d.data.location + '<br>' : ''}
        ${d.data.email ? d.data.email + '<br>' : ''}
        Total Reports: ${totalReports}<br>
        Licensed Users: ${licensedReports} (${licensePercent}%)<br>
        <div style="margin-top: 0.5rem; padding-top: 0.5rem; border-top: 1px solid #ddd;">
          ${licenseText}
        </div>
        ${engagementText}
      `;
      
      tooltip.style.left = (event.pageX + 10) + 'px';
      tooltip.style.top = (event.pageY - 10) + 'px';
      tooltip.classList.add('visible');
    },
    
    // Hide tooltip
    hideTooltip: function() {
      document.getElementById('tooltip').classList.remove('visible');
    },

    // Update selected user panel
    updateSelectedUser: function(d) {
      const section = document.getElementById('selectedUserSection');
      if (!section) return;

      const totalReports = countDescendants(d);
      const licensedReports = countLicensedDescendants(d);
      const licensePercent = totalReports > 0
        ? Number(((licensedReports / totalReports) * 100).toFixed(1))
        : 0;
      const emailHtml = d.data.email ? `
        <div class="selected-user-email">
          <span>${d.data.email}</span>
          <button class="copy-btn">Copy</button>
        </div>` : '';

      section.innerHTML = `
        <h4>Selected User</h4>
        <div><strong>${d.data.name || 'Unknown'}</strong></div>
        ${d.data.title ? `<div>${d.data.title}</div>` : ''}
        <div>${d.data.department || 'No department'}</div>
        ${d.data.location ? `<div>${d.data.location}</div>` : ''}
        ${emailHtml}
        <div>Total Reports: ${totalReports}</div>
        <div>Licensed Users: ${licensedReports} (${licensePercent}%)</div>
      `;
      section.style.display = 'block';

      if (d.data.email) {
        const btn = section.querySelector('.copy-btn');
        if (btn) {
          btn.addEventListener('click', () => navigator.clipboard.writeText(d.data.email));
        }
      }
    },
    
    // Update legend based on color scheme
    updateLegend: function() {
      const legendContent = document.getElementById('legendContent');
      const colorBy = document.getElementById('colorBy').value;

      let html = '<h4>Legend</h4>';
      
      // Always show license border legend
      html += `
        <div style="margin-bottom: 1rem; padding-bottom: 0.5rem; border-bottom: 1px solid var(--border);">
          <strong>License Status (Border)</strong>
          <div class="legend-item">
            <div class="legend-color" style="background: white; border: 3px solid #00a854;"></div>
            <span>Has ChatGPT License</span>
          </div>
          <div class="legend-item">
            <div class="legend-color" style="background: white; border: 3px solid #dc3545;"></div>
            <span>No License</span>
          </div>
        </div>
      `;
      
      // Add color-specific legend
      html += '<strong>Node Color</strong>';
      
      if (colorBy === 'license') {
        html += `
          <div class="legend-item">
            <div class="legend-color" style="background: #00a854;"></div>
            <span>Has License (Fill)</span>
          </div>
          <div class="legend-item">
            <div class="legend-color" style="background: #f0f0f0;"></div>
            <span>No License (Fill)</span>
          </div>
        `;
      } else if (colorBy === 'level') {
        for (let i = 0; i <= 4; i++) {
          html += `
            <div class="legend-item">
              <div class="legend-color" style="background: ${levelColorScale(i)};"></div>
              <span>Level ${i}</span>
            </div>
          `;
        }
      } else { // department
        const departments = [...new Set(orgData.map(u => u.department))].slice(0, 10);
        departments.forEach(dept => {
          html += `
            <div class="legend-item">
              <div class="legend-color" style="background: ${departmentColorScale(dept)};"></div>
              <span>${dept}</span>
            </div>
          `;
        });
        if (departments.length < [...new Set(orgData.map(u => u.department))].length) {
          html += '<div class="legend-item"><small>...and more</small></div>';
        }
      }
      
      legendContent.innerHTML = html;
    },
    
    // Update statistics panel
    updateStatsPanel: function() {
      const stats = document.getElementById('statsPanel');
      const totalUsers = orgData.length;
      const licensed = orgData.filter(u => u.hasLicense).length;
      const licensedPercent = totalUsers > 0
        ? Math.round((licensed / totalUsers) * 100)
        : 0;
      const departments = [...new Set(orgData.map(u => u.department))].length;
      const visibleNodes = root ? root.descendants() : [];
      const visibleLicensed = visibleNodes.filter(d => d.data.hasLicense).length;
      const visibleLicensedPercent = visibleNodes.length > 0
        ? Math.round((visibleLicensed / visibleNodes.length) * 100)
        : 0;
      const highlightedVisibleNodes = highlightedNodes
        ? visibleNodes.filter(d => highlightedNodes.has(d.data.email))
        : [];
      const isolatedCount = highlightedVisibleNodes.length;
      const isolatedLicensed = highlightedVisibleNodes.filter(d => d.data.hasLicense).length;
      const isolatedLicensedPercent = isolatedCount > 0
        ? Math.round((isolatedLicensed / isolatedCount) * 100)
        : 0;

      stats.innerHTML = `
        <div class="stat-item">
          <span class="stat-label">Total Users:</span>
          <span class="stat-value">${totalUsers}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Licensed:</span>
          <span class="stat-value">${licensed} (${licensedPercent}%)</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Departments:</span>
          <span class="stat-value">${departments}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Isolated Users:</span>
          <span class="stat-value">${isolatedCount}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Isolated Licensed Users:</span>
          <span class="stat-value">${isolatedLicensed} (${isolatedLicensedPercent}%)</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Visible Users:</span>
          <span class="stat-value">${visibleNodes.length}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Visible Licensed Users:</span>
          <span class="stat-value">${visibleLicensed} (${visibleLicensedPercent}%)</span>
        </div>
      `;
    }
  };
})();

// Expose OrgChart to global scope for inline event handlers
window.OrgChart = OrgChart;