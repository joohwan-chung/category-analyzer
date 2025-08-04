class CategoryAnalyzer {
  constructor() {
    this.data = null;
    this.treeData = null;
    this.selectedItem = null;
    this.searchTerm = '';
    this.isProcessing = false; // 진행 상태 플래그 추가
    this.isMultiSelectMode = false; // 다중 선택 모드 플래그
    this.selectedItems = []; // 다중 선택된 아이템들
    this.multiInputRowCount = 0; // 다중 입력 행 카운터

    this.initializeElements();
    this.bindEvents();
    
    // CategoryDownloader와 연동
    if (window.categoryDownloader) {
      window.categoryDownloader.setAnalyzer(this);
    }
  }

  initializeElements() {
    this.fileInput = document.getElementById('excelFile');
    this.analyzeBtn = document.getElementById('analyzeBtn');
    this.addCategoryBtn = document.getElementById('addCategoryBtn');
    this.multiInputBtn = document.getElementById('multiInputBtn');
    this.loading = document.getElementById('loading');
    this.analysisSection = document.querySelector('.analysis-section');
    this.inputSection = document.querySelector('.input-section');
    this.editSection = document.querySelector('.edit-section');
    this.multiInputSection = document.querySelector('.multi-input-section');
    this.multiEditSection = document.querySelector('.multi-edit-section');
    this.treeView = document.getElementById('treeView');
    this.categoryDetails = document.getElementById('categoryDetails');
    this.searchInput = document.getElementById('searchInput');
    this.expandAllBtn = document.getElementById('expandAllBtn');
    this.collapseAllBtn = document.getElementById('collapseAllBtn');
    this.clearSelectionBtn = document.getElementById('clearSelectionBtn');
    this.editCategoryBtn = document.getElementById('editCategoryBtn');
    this.deleteCategoryBtn = document.getElementById('deleteCategoryBtn');
    this.multiSelectModeBtn = document.getElementById('multiSelectModeBtn');
    this.multiEditBtn = document.getElementById('multiEditBtn');
    this.multiDeleteBtn = document.getElementById('multiDeleteBtn');
    this.categoryForm = document.getElementById('categoryForm');
    this.editCategoryForm = document.getElementById('editCategoryForm');
    this.multiCategoryForm = document.getElementById('multiCategoryForm');
    this.multiEditForm = document.getElementById('multiEditForm');
    this.cancelBtn = document.getElementById('cancelBtn');
    this.editCancelBtn = document.getElementById('editCancelBtn');
    this.multiInputCancelBtn = document.getElementById('multiInputCancelBtn');
    this.multiEditCancelBtn = document.getElementById('multiEditCancelBtn');
    this.addRowBtn = document.getElementById('addRowBtn');
    this.removeRowBtn = document.getElementById('removeRowBtn');
    this.clearAllRowsBtn = document.getElementById('clearAllRowsBtn');
    this.selectAllEditBtn = document.getElementById('selectAllEditBtn');
    this.deselectAllEditBtn = document.getElementById('deselectAllEditBtn');
    this.multiFormRows = document.getElementById('multiFormRows');
    this.multiEditRows = document.getElementById('multiEditRows');

    // 통계 요소들
    this.totalCategoriesEl = document.getElementById('totalCategories');
    this.maxDepthEl = document.getElementById('maxDepth');
    this.domainTypesEl = document.getElementById('domainTypes');
    this.rootCategoriesEl = document.getElementById('rootCategories');
    this.childCategoriesEl = document.getElementById('childCategories');
    this.hasDomainCountEl = document.getElementById('hasDomainCount');
    this.noDomainCountEl = document.getElementById('noDomainCount');
    this.avgDepthEl = document.getElementById('avgDepth');
    this.domainTypeStatsEl = document.getElementById('domainTypeStats');
    this.depthDistributionEl = document.getElementById('depthDistribution');
    
      // 분석 통계 펼쳐보기/접기 요소들
  this.statsToggleBtn = document.getElementById('statsToggleBtn');
  this.statsContent = document.getElementById('statsContent');
  }

  bindEvents() {
    this.fileInput.addEventListener('change', this.handleFileSelect.bind(this));
    this.analyzeBtn.addEventListener('click', this.analyzeData.bind(this));
    this.addCategoryBtn.addEventListener('click', this.showInputSection.bind(this));
    this.multiInputBtn.addEventListener('click', this.showMultiInputSection.bind(this));
    this.searchInput.addEventListener('input', (e) => {
      this.filterTree();
    });
    this.expandAllBtn.addEventListener('click', this.expandAll.bind(this));
    this.collapseAllBtn.addEventListener('click', this.collapseAll.bind(this));
    this.clearSelectionBtn.addEventListener('click', this.clearSelection.bind(this));
    this.editCategoryBtn.addEventListener('click', this.showEditSection.bind(this));
    this.deleteCategoryBtn.addEventListener('click', this.deleteCategory.bind(this));
    this.multiSelectModeBtn.addEventListener('click', this.toggleMultiSelectMode.bind(this));
    this.multiEditBtn.addEventListener('click', this.showMultiEditSection.bind(this));
    this.multiDeleteBtn.addEventListener('click', this.multiDeleteCategories.bind(this));
    this.categoryForm.addEventListener('submit', this.addNewCategory.bind(this));
    this.editCategoryForm.addEventListener('submit', this.updateCategory.bind(this));
    this.multiCategoryForm.addEventListener('submit', this.addMultipleCategories.bind(this));
    this.multiEditForm.addEventListener('submit', this.updateMultipleCategories.bind(this));
    this.cancelBtn.addEventListener('click', this.hideInputSection.bind(this));
    this.editCancelBtn.addEventListener('click', this.hideEditSection.bind(this));
    this.multiInputCancelBtn.addEventListener('click', this.hideMultiInputSection.bind(this));
    this.multiEditCancelBtn.addEventListener('click', this.hideMultiEditSection.bind(this));
    this.addRowBtn.addEventListener('click', this.addMultiInputRow.bind(this));
    this.removeRowBtn.addEventListener('click', this.removeSelectedRows.bind(this));
    this.clearAllRowsBtn.addEventListener('click', this.clearAllMultiInputRows.bind(this));
    this.selectAllEditBtn.addEventListener('click', this.selectAllEditRows.bind(this));
    this.deselectAllEditBtn.addEventListener('click', this.deselectAllEditRows.bind(this));
    
    // 분석 통계 토글 버튼 이벤트 바인딩
    if (this.statsToggleBtn) {
      this.statsToggleBtn.addEventListener('click', this.toggleStatsSection.bind(this));
    }

    // 트리 뷰에서 빈 공간 클릭 시 선택 해제
    this.treeView.addEventListener('click', (e) => {
      if (e.target === this.treeView || e.target.classList.contains('tree-view')) {
        this.clearSelection();
      }
    });
  }

  handleFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
      this.analyzeBtn.disabled = false;
      this.analyzeBtn.title = "클릭하여 엑셀 파일을 분석하세요";
      // 새 카테고리 추가 버튼은 분석 완료 후에만 활성화
      this.addCategoryBtn.disabled = true;
      this.addCategoryBtn.title = "먼저 엑셀 파일을 분석한 후 사용할 수 있습니다";
      this.multiInputBtn.disabled = true;
      this.multiInputBtn.title = "먼저 엑셀 파일을 분석한 후 사용할 수 있습니다";
      
      // 다운로드 버튼 비활성화
      if (window.categoryDownloader) {
        window.categoryDownloader.disableDownloadButton();
      }
      
      this.fileInput.files = event.target.files;
    } else {
      // 파일이 선택되지 않은 경우 버튼 비활성화
      this.analyzeBtn.disabled = true;
      this.analyzeBtn.title = "엑셀 파일을 선택한 후 클릭하세요";
      this.addCategoryBtn.disabled = true;
      this.addCategoryBtn.title = "먼저 엑셀 파일을 분석한 후 사용할 수 있습니다";
      this.multiInputBtn.disabled = true;
      this.multiInputBtn.title = "먼저 엑셀 파일을 분석한 후 사용할 수 있습니다";
      
      // 다운로드 버튼 비활성화
      if (window.categoryDownloader) {
        window.categoryDownloader.disableDownloadButton();
      }
    }
  }

  async analyzeData() {
    const file = this.fileInput.files[0];
    if (!file) return;

    this.showLoading(true);

    try {
      const data = await this.readExcelFile(file);
      this.data = this.processData(data);
      this.treeData = this.buildTreeStructure();

      this.updateStatistics();
      this.renderTree();
      this.showAnalysisSection();

      // 분석 완료 후 새 카테고리 추가 버튼 활성화
      this.addCategoryBtn.disabled = false;
      this.addCategoryBtn.title = "클릭하여 새 카테고리를 추가하세요";
      this.multiInputBtn.disabled = false;
      this.multiInputBtn.title = "클릭하여 여러 카테고리를 한 번에 추가하세요";
      
      // 다운로드 버튼 활성화
      if (window.categoryDownloader) {
        window.categoryDownloader.enableDownloadButton();
      }

    } catch (error) {
      console.error('데이터 분석 중 오류 발생:', error);
      alert('파일을 읽는 중 오류가 발생했습니다. 파일 형식을 확인해주세요.');
    } finally {
      this.showLoading(false);
    }
  }

  readExcelFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          resolve(jsonData);
        } catch (error) {
          reject(error);
        }
      };

      reader.onerror = () => reject(new Error('파일 읽기 실패'));
      reader.readAsArrayBuffer(file);
    });
  }

  processData(rawData) {
    if (!rawData || rawData.length < 2) {
      throw new Error('유효하지 않은 데이터입니다.');
    }

    const headers = rawData[0];
    const rows = rawData.slice(1);

    // 헤더 매핑
    const headerMap = {
      '일련번호': 'id',
      '부모 카테고리번호': 'parentId',
      '부모 카테고리명': 'parentName',
      '도메인구분': 'domainType',
      '도메인주소여부': 'hasDomain',
      '도메인주소': 'domainAddress'
    };

    // 메모리 효율성을 위해 필요한 컬럼만 처리
    const processedData = [];
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const item = {};

      headers.forEach((header, colIndex) => {
        const mappedKey = headerMap[header];
        if (mappedKey) {
          let value = row[colIndex] || '';

          // id와 parentId는 문자열로 통일
          if (mappedKey === 'id' || mappedKey === 'parentId') {
            value = value.toString();
          }

          item[mappedKey] = value;
        }
      });
      item.originalIndex = i;
      processedData.push(item);
    }

    // '기본 분석 목록'이 없으면 자동 추가
    let basicRoot = processedData.find(item => item.parentName === '기본 분석 목록');
    if (!basicRoot) {
      const rootId = 'ROOT001';
      basicRoot = {
        id: rootId,
        parentId: '',
        parentName: '기본 분석 목록',
        domainType: '',
        hasDomain: '',
        domainAddress: '',
        originalIndex: -1
      };
      processedData.unshift(basicRoot);
    }
    this.basicRootId = basicRoot.id;

    return processedData;
  }

  buildTreeStructure() {
    const treeMap = new Map();
    const rootNodes = [];

    // 모든 노드를 맵에 추가
    this.data.forEach(item => {
      treeMap.set(item.id, {
        ...item,
        children: [],
        level: 0
      });
    });

    // 부모-자식 관계 설정
    this.data.forEach(item => {
      const node = treeMap.get(item.id);
      if (item.parentId && item.parentId !== '' && item.parentId !== '0') {
        const parent = treeMap.get(item.parentId);
        if (parent) {
          parent.children.push(node);
          node.level = parent.level + 1;
        } else {
          // 부모가 존재하지 않는 경우 루트 노드로 처리
          rootNodes.push(node);
        }
      } else {
        // parentId가 비어있거나 0인 경우 루트 노드로 처리
        rootNodes.push(node);
      }
    });

    return rootNodes;
  }

  // 아이템의 깊이 레벨을 찾는 헬퍼 메서드
  getItemLevel(itemId) {
    const findLevel = (nodes, targetId, currentLevel = 0) => {
      for (const node of nodes) {
        if (node.id === targetId) {
          return currentLevel;
        }
        if (node.children && node.children.length > 0) {
          const found = findLevel(node.children, targetId, currentLevel + 1);
          if (found !== null) {
            return found;
          }
        }
      }
      return null;
    };

    return findLevel(this.treeData, itemId);
  }

  updateStatistics() {
    const totalCategories = this.data.length;
    const maxDepth = this.calculateMaxDepth(this.treeData);
    const domainTypes = new Set(this.data.map(item => item.domainType).filter(type => type));

    // 루트 카테고리와 하위 카테고리 계산
    const rootCategories = this.data.filter(item => !item.parentId || item.parentId === '').length;
    const childCategories = totalCategories - rootCategories;

    // 도메인 주소 보유 여부 통계
    // 도메인 구분이 'D'이면 도메인 주소가 있다고 판단
    const hasDomainCount = this.data.filter(item => 
      item.hasDomain === 'Y' || item.domainType === 'D'
    ).length;
    
    // 도메인 구분이 'H'인 항목들을 도메인 주소 미보유로 카운트
    const noDomainCount = this.data.filter(item => 
      item.domainType === 'H'
    ).length;

    // 평균 깊이 계산
    const avgDepth = this.calculateAverageDepth(this.treeData);

    // 도메인 구분별 통계
    const domainTypeStats = this.calculateDomainTypeStats();

    // 깊이별 분포 계산
    const depthDistribution = this.calculateDepthDistribution();

    // 기본 통계 업데이트
    this.totalCategoriesEl.textContent = totalCategories;
    this.maxDepthEl.textContent = maxDepth;
    this.domainTypesEl.textContent = domainTypes.size;
    this.rootCategoriesEl.textContent = rootCategories;
    this.childCategoriesEl.textContent = childCategories;
    this.hasDomainCountEl.textContent = hasDomainCount;
    this.noDomainCountEl.textContent = noDomainCount;
    this.avgDepthEl.textContent = avgDepth.toFixed(1);

    // 상세 통계 업데이트
    this.updateDomainTypeStats(domainTypeStats);
    this.updateDepthDistribution(depthDistribution);
  }

  calculateMaxDepth(nodes, currentDepth = 0) {
    let maxDepth = currentDepth;
    nodes.forEach(node => {
      if (node.children && node.children.length > 0) {
        maxDepth = Math.max(maxDepth, this.calculateMaxDepth(node.children, currentDepth + 1));
      }
    });
    return maxDepth;
  }

  calculateAverageDepth(nodes, currentDepth = 0) {
    let totalDepth = 0;
    let nodeCount = 0;

    const calculateDepth = (nodes, depth) => {
      nodes.forEach(node => {
        totalDepth += depth;
        nodeCount++;
        if (node.children && node.children.length > 0) {
          calculateDepth(node.children, depth + 1);
        }
      });
    };

    calculateDepth(nodes, currentDepth);
    return nodeCount > 0 ? totalDepth / nodeCount : 0;
  }

  calculateDomainTypeStats() {
    const domainTypeMap = {};
    this.data.forEach(item => {
      if (item.domainType) {
        if (!domainTypeMap[item.domainType]) {
          domainTypeMap[item.domainType] = {
            count: 0,
            hasDomain: 0,
            noDomain: 0
          };
        }
        domainTypeMap[item.domainType].count++;
        
        // 도메인 구분이 'D'이면 도메인 주소가 있다고 판단
        if (item.domainType === 'D' || item.hasDomain === 'Y') {
          domainTypeMap[item.domainType].hasDomain++;
        } else if (item.hasDomain === 'N' || !item.hasDomain || item.hasDomain === '') {
          domainTypeMap[item.domainType].noDomain++;
        }
      }
    });
    return domainTypeMap;
  }

  calculateDepthDistribution() {
    const depthMap = {};

    const calculateDepth = (nodes, depth) => {
      nodes.forEach(node => {
        if (!depthMap[depth]) {
          depthMap[depth] = 0;
        }
        depthMap[depth]++;

        if (node.children && node.children.length > 0) {
          calculateDepth(node.children, depth + 1);
        }
      });
    };

    calculateDepth(this.treeData, 0);
    return depthMap;
  }

  updateDomainTypeStats(domainTypeStats) {
    this.domainTypeStatsEl.innerHTML = '';

    if (Object.keys(domainTypeStats).length === 0) {
      this.domainTypeStatsEl.innerHTML = '<p class="no-data">도메인 구분 데이터가 없습니다.</p>';
      return;
    }

    Object.entries(domainTypeStats)
      .sort(([, a], [, b]) => b.count - a.count) // 개수 기준 내림차순 정렬
      .forEach(([domainType, stats]) => {
        const statItem = document.createElement('div');
        statItem.className = 'detail-item';
        
        // 도메인 구분에 따른 표시
        let domainTypeLabel = domainType;
        if (domainType === 'D') {
          domainTypeLabel = `${domainType} (도메인 주소 보유)`;
        } else if (domainType === 'H') {
          domainTypeLabel = `${domainType} (도메인 주소 미보유)`;
        }
        
        statItem.innerHTML = `
          <span class="detail-label">${domainTypeLabel}</span>
          <span class="detail-value">${stats.count}개</span>
        `;
        this.domainTypeStatsEl.appendChild(statItem);
      });
  }

  updateDepthDistribution(depthDistribution) {
    this.depthDistributionEl.innerHTML = '';

    if (Object.keys(depthDistribution).length === 0) {
      this.depthDistributionEl.innerHTML = '<p class="no-data">깊이 분포 데이터가 없습니다.</p>';
      return;
    }

    Object.entries(depthDistribution)
      .sort(([a], [b]) => parseInt(a) - parseInt(b)) // 깊이 기준 오름차순 정렬
      .forEach(([depth, count]) => {
        const statItem = document.createElement('div');
        statItem.className = 'detail-item';
        statItem.innerHTML = `
          <span class="detail-label">깊이 ${depth}</span>
          <span class="detail-value">${count}개</span>
        `;
        this.depthDistributionEl.appendChild(statItem);
      });
  }

  renderTree() {
    this.treeView.innerHTML = '';
    this.renderTreeNodes(this.treeData, this.treeView);
  }

  renderTreeNodes(nodes, container) {
    nodes.forEach(node => {
      const treeItem = this.createTreeNode(node);
      container.appendChild(treeItem);
    });
  }

  createTreeNode(node) {
    const treeItem = document.createElement('div');
    treeItem.className = 'tree-item';
    treeItem.dataset.id = node.id;

    const treeContent = document.createElement('div');
    treeContent.className = 'tree-content';

    // 다중 선택 모드일 때 체크박스 추가
    if (this.isMultiSelectMode) {
      treeItem.classList.add('multi-select-mode');
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.className = 'tree-checkbox';
      checkbox.addEventListener('change', (e) => {
        e.stopPropagation();
        this.handleMultiSelect(treeItem, node, checkbox.checked);
      });
      treeContent.appendChild(checkbox);
    }

    const toggle = document.createElement('span');
    toggle.className = 'tree-toggle';
    toggle.textContent = '▶';
    toggle.style.display = node.children.length > 0 ? 'inline-block' : 'none';

    const icon = document.createElement('span');
    icon.className = 'tree-icon';
    icon.textContent = node.children.length > 0 ? '📁' : '📄';

    const label = document.createElement('span');
    label.className = 'tree-label';
    label.textContent = node.parentName || `카테고리 ${node.id}`;

    // 새로 추가된 카테고리인지 확인 (originalIndex가 있는 경우)
    if (node.originalIndex !== undefined) {
      treeItem.classList.add('new-category');
      label.style.color = '#007bff'; // 파란색으로 표시
    }

    treeContent.appendChild(toggle);
    treeContent.appendChild(icon);
    treeContent.appendChild(label);

    treeItem.appendChild(treeContent);

    if (node.children.length > 0) {
      const childrenContainer = document.createElement('div');
      childrenContainer.className = 'tree-children';
      this.renderTreeNodes(node.children, childrenContainer);
      treeItem.appendChild(childrenContainer);

      toggle.addEventListener('click', (e) => {
        e.stopPropagation();
        this.toggleNode(treeItem, toggle, childrenContainer);
      });
    }

    treeContent.addEventListener('click', () => {
      if (!this.isMultiSelectMode) {
        this.selectNode(treeItem, node);
      }
    });

    return treeItem;
  }

  toggleNode(treeItem, toggle, childrenContainer) {
    const isExpanded = childrenContainer.classList.contains('expanded');

    if (isExpanded) {
      // 접기 - 부드러운 애니메이션
      childrenContainer.style.opacity = '0';
      childrenContainer.style.transform = 'translateY(-10px)';
      
      setTimeout(() => {
        childrenContainer.classList.remove('expanded');
        toggle.classList.remove('expanded');
        childrenContainer.style.display = 'none';
      }, 300); // CSS transition 시간과 동일
    } else {
      // 펼치기 - 부드러운 애니메이션
      childrenContainer.classList.add('expanded');
      toggle.classList.add('expanded');
      childrenContainer.style.display = 'block';
      
      setTimeout(() => {
        childrenContainer.style.opacity = '1';
        childrenContainer.style.transform = 'translateY(0)';
      }, 10);
    }
  }

  selectNode(treeItem, node) {
    // 이전 선택 해제
    document.querySelectorAll('.tree-item.selected').forEach(item => {
      item.classList.remove('selected');
    });

    // 새 선택
    treeItem.classList.add('selected');
    this.selectedItem = node;
    this.showCategoryDetails(node);

    // 하위 카테고리 추가 버튼 활성화 및 상태 업데이트
    this.addCategoryBtn.disabled = false;
    this.addCategoryBtn.title = `"${node.parentName}" 하위에 새 카테고리를 추가합니다`;

    // 버튼 텍스트 업데이트
    this.addCategoryBtn.textContent = `"${node.parentName}" 하위 추가`;

    // 수정/삭제 버튼 활성화
    this.editCategoryBtn.disabled = false;
    this.editCategoryBtn.title = `"${node.parentName}" 카테고리를 수정합니다`;
    this.deleteCategoryBtn.disabled = false;
    this.deleteCategoryBtn.title = `"${node.parentName}" 카테고리를 삭제합니다`;

    // 선택된 카테고리 정보를 화면에 표시
    this.showSelectedCategoryInfo(node);
  }

  // 트리에서 선택 해제하는 함수 추가
  clearSelection() {
    // 선택 해제
    document.querySelectorAll('.tree-item.selected').forEach(item => {
      item.classList.remove('selected');
    });

    this.selectedItem = null;

    // 하위 카테고리 추가 버튼 비활성화
    this.addCategoryBtn.disabled = true;
    this.addCategoryBtn.textContent = '하위 카테고리 추가';
    this.addCategoryBtn.title = '트리에서 부모 카테고리를 선택한 후 사용할 수 있습니다';

    // 수정/삭제 버튼 비활성화
    this.editCategoryBtn.disabled = true;
    this.editCategoryBtn.title = '트리에서 카테고리를 선택한 후 수정할 수 있습니다';
    this.deleteCategoryBtn.disabled = true;
    this.deleteCategoryBtn.title = '트리에서 카테고리를 선택한 후 삭제할 수 있습니다';

    // 선택된 카테고리 정보 숨기기
    const infoElement = document.getElementById('selectedCategoryInfo');
    if (infoElement) {
      infoElement.remove();
    }
  }

  showCategoryDetails(node) {
    const isNewCategory = node.originalIndex !== undefined;
    const statusText = isNewCategory ? '🆕 새로 추가된 카테고리 (기본 분석 목록 포함)' : '기존 카테고리';

    const details = `
            <div class="detail-item">
                <span class="detail-label">상태</span>
                <span class="detail-value" style="color: ${isNewCategory ? '#27ae60' : '#2c3e50'}; font-weight: ${isNewCategory ? '600' : 'normal'};">${statusText}</span>
            </div>
            <div class="detail-item">
                <span class="detail-label">일련번호</span>
                <span class="detail-value">${node.id || 'N/A'}</span>
            </div>
            <div class="detail-item">
                <span class="detail-label">부모 카테고리명</span>
                <span class="detail-value">${node.parentName || 'N/A'}</span>
            </div>
            <div class="detail-item">
                <span class="detail-label">도메인 구분</span>
                <span class="detail-value">${node.domainType || 'N/A'}</span>
            </div>
            <div class="detail-item">
                <span class="detail-label">도메인 주소 여부</span>
                <span class="detail-value">${node.hasDomain || 'N/A'}</span>
            </div>
            <div class="detail-item">
                <span class="detail-label">도메인 주소</span>
                <span class="detail-value">${node.domainAddress || 'N/A'}</span>
            </div>
            <div class="detail-item">
                <span class="detail-label">하위 카테고리 수</span>
                <span class="detail-value">${node.children.length}</span>
            </div>
            <div class="detail-item">
                <span class="detail-label">트리 레벨</span>
                <span class="detail-value">${node.level}</span>
            </div>
        `;

    this.categoryDetails.innerHTML = details;
  }

  filterTree() {
    const searchTerm = this.searchInput.value.trim();

    if (!searchTerm) {
      this.showAllNodes();
      this.clearAllHighlights();
      this.hideNoResultsMessage();
      return;
    }

    this.hideAllNodes();
    this.clearAllHighlights();
    this.showMatchingNodes(this.treeData, searchTerm.toLowerCase());

    // 검색 결과가 있는지 확인
    const visibleItems = document.querySelectorAll('.tree-item[style*="display: block"]');
    if (visibleItems.length === 0) {
      this.showNoResultsMessage(searchTerm);
    } else {
      this.hideNoResultsMessage();
    }
  }

  showAllNodes() {
    document.querySelectorAll('.tree-item').forEach(item => {
      item.style.display = 'block';
    });
  }

  hideAllNodes() {
    document.querySelectorAll('.tree-item').forEach(item => {
      item.style.display = 'none';
    });
  }

  clearAllHighlights() {
    document.querySelectorAll('.highlight').forEach(highlight => {
      const parent = highlight.parentNode;
      if (parent) {
        parent.innerHTML = parent.innerHTML.replace(/<span class="highlight">(.*?)<\/span>/g, '$1');
      }
    });
  }

  showMatchingNodes(nodes, searchTerm) {
    nodes.forEach(node => {
      const treeItem = document.querySelector(`[data-id="${node.id}"]`);
      if (!treeItem) return;

      const nodeText = (node.parentName || '').toLowerCase();
      const matches = nodeText.includes(searchTerm);

      if (matches) {
        treeItem.style.display = 'block';
        this.highlightSearchTerm(treeItem, searchTerm);
        this.showParentNodes(treeItem);
      }

      // 자식 노드들도 검색
      if (node.children.length > 0) {
        this.showMatchingNodes(node.children, searchTerm);
      }
    });
  }

  highlightSearchTerm(treeItem, searchTerm) {
    const label = treeItem.querySelector('.tree-label');
    if (!label) return;

    const text = label.textContent;
    const regex = new RegExp(`(${searchTerm})`, 'gi');
    const highlightedText = text.replace(regex, '<span class="highlight">$1</span>');
    label.innerHTML = highlightedText;
  }

  showParentNodes(treeItem) {
    let parent = treeItem.parentElement.closest('.tree-item');
    while (parent) {
      parent.style.display = 'block';
      const childrenContainer = parent.querySelector('.tree-children');
      if (childrenContainer) {
        childrenContainer.classList.add('expanded');
        childrenContainer.style.display = 'block';
        const toggle = parent.querySelector('.tree-toggle');
        if (toggle) {
          toggle.classList.add('expanded');
        }
      }
      parent = parent.parentElement.closest('.tree-item');
    }
  }

  async expandAll() {
    if (this.isProcessing) return; // 이미 처리 중이면 중단

    this.isProcessing = true;
    this.expandAllBtn.disabled = true;
    this.collapseAllBtn.disabled = true;

    // 접힌 노드들만 찾기
    const collapsedNodes = document.querySelectorAll('.tree-children:not(.expanded)');
    const totalNodes = collapsedNodes.length;
    let processedNodes = 0;

    // 진행 상태 표시 UI 생성
    const progressContainer = this.createProgressUI();

    // 배치 처리로 진행률 업데이트
    const batchSize = 10;
    const processBatch = () => {
      const startIndex = processedNodes;
      const endIndex = Math.min(startIndex + batchSize, totalNodes);

      for (let i = startIndex; i < endIndex; i++) {
        const container = collapsedNodes[i];
        const toggle = container.parentElement.querySelector('.tree-toggle');

        container.classList.add('expanded');
        // 인라인 스타일을 명시적으로 설정하여 확실히 펼쳐지도록 함
        container.style.display = 'block';
        if (toggle) {
          toggle.classList.add('expanded');
        }
        processedNodes++;
      }

      // 진행률 업데이트
      const progress = (processedNodes / totalNodes) * 100;
      this.updateProgress(progressContainer, progress, processedNodes, totalNodes);

      // 다음 배치 처리 또는 완료
      if (processedNodes < totalNodes) {
        setTimeout(processBatch, 10); // 10ms 지연으로 UI 업데이트 시간 확보
      } else {
        setTimeout(() => {
          this.hideProgressUI(progressContainer);
          this.isProcessing = false;
          this.expandAllBtn.disabled = false;
          this.collapseAllBtn.disabled = false;
        }, 500); // 완료 후 0.5초 대기
      }
    };

    processBatch();
  }

  async collapseAll() {
    if (this.isProcessing) return; // 이미 처리 중이면 중단

    this.isProcessing = true;
    this.expandAllBtn.disabled = true;
    this.collapseAllBtn.disabled = true;

    // 펼쳐진 노드들만 찾기
    const expandedNodes = document.querySelectorAll('.tree-children.expanded');
    const totalNodes = expandedNodes.length;
    let processedNodes = 0;

    // 진행 상태 표시 UI 생성
    const progressContainer = this.createProgressUI('접는 중...');

    // 배치 처리로 진행률 업데이트
    const batchSize = 10;
    const processBatch = () => {
      const startIndex = processedNodes;
      const endIndex = Math.min(startIndex + batchSize, totalNodes);

      for (let i = startIndex; i < endIndex; i++) {
        const container = expandedNodes[i];
        const toggle = container.parentElement.querySelector('.tree-toggle');

        container.classList.remove('expanded');
        // 인라인 스타일도 제거하여 확실히 접히도록 함
        container.style.display = '';
        if (toggle) {
          toggle.classList.remove('expanded');
        }
        processedNodes++;
      }

      // 진행률 업데이트
      const progress = (processedNodes / totalNodes) * 100;
      this.updateProgress(progressContainer, progress, processedNodes, totalNodes);

      // 다음 배치 처리 또는 완료
      if (processedNodes < totalNodes) {
        setTimeout(processBatch, 10);
      } else {
        setTimeout(() => {
          this.hideProgressUI(progressContainer);
          this.isProcessing = false;
          this.expandAllBtn.disabled = false;
          this.collapseAllBtn.disabled = false;
        }, 500);
      }
    };

    processBatch();
  }

  showLoading(show) {
    this.loading.style.display = show ? 'flex' : 'none';
  }

  showAnalysisSection() {
    this.analysisSection.style.display = 'grid';
  }

  async showInputSection() {
    // 분석이 완료되지 않았으면 경고 메시지 표시
    if (!this.data || this.data.length === 0) {
      alert('먼저 엑셀 파일을 분석해주세요.');
      return;
    }

    // 선택된 카테고리가 없으면 경고
    if (!this.selectedItem) {
      alert('트리에서 부모 카테고리를 먼저 선택해주세요.');
      return;
    }

    this.inputSection.style.display = 'block';
    this.analysisSection.style.display = 'none';
    this.clearForm();

    // 일련번호 자동 채우기
    this.autoFillNextId();

    // 선택된 카테고리를 부모로 자동 설정
    const parentInput = document.getElementById('newParentId');
    if (parentInput && this.selectedItem) {
      parentInput.value = `${this.selectedItem.id} - ${this.selectedItem.parentName}`;
      parentInput.dataset.selectedValue = this.selectedItem.id;
    }

    // 드롭다운 채우기를 비동기로 처리
    try {
      await this.populateDropdowns();
    } catch (error) {
      console.error('드롭다운 채우기 중 오류:', error);
      alert('드롭다운 데이터 로딩 중 오류가 발생했습니다.');
    }
  }

  hideInputSection() {
    this.inputSection.style.display = 'none';
    this.analysisSection.style.display = 'grid';
  }

  async showEditSection() {
    // 분석이 완료되지 않았으면 경고 메시지 표시
    if (!this.data || this.data.length === 0) {
      alert('먼저 엑셀 파일을 분석해주세요.');
      return;
    }

    // 선택된 카테고리가 없으면 경고
    if (!this.selectedItem) {
      alert('트리에서 수정할 카테고리를 먼저 선택해주세요.');
      return;
    }

    // 기본 분석 목록은 수정 불가
    if (this.selectedItem.parentName === '기본 분석 목록') {
      alert('기본 분석 목록은 수정할 수 없습니다.');
      return;
    }

    this.editSection.style.display = 'block';
    this.analysisSection.style.display = 'none';
    this.clearEditForm();

    // 선택된 카테고리 정보로 폼 채우기
    this.fillEditForm(this.selectedItem);

    // 드롭다운 채우기를 비동기로 처리
    try {
      await this.populateEditDropdowns();
    } catch (error) {
      console.error('드롭다운 채우기 중 오류:', error);
      alert('드롭다운 데이터 로딩 중 오류가 발생했습니다.');
    }
  }

  hideEditSection() {
    this.editSection.style.display = 'none';
    this.analysisSection.style.display = 'grid';
  }

  clearEditForm() {
    this.editCategoryForm.reset();
  }

  fillEditForm(category) {
    document.getElementById('editId').value = category.id;
    document.getElementById('editParentName').value = category.parentName;
    document.getElementById('editDomainType').value = category.domainType || '';
    document.getElementById('editHasDomain').value = category.hasDomain || '';
    document.getElementById('editDomainAddress').value = category.domainAddress || '';

    // 부모 카테고리 설정
    const parentInput = document.getElementById('editParentId');
    if (category.parentId && category.parentId !== '') {
      const parentCategory = this.data.find(item => item.id === category.parentId);
      if (parentCategory) {
        parentInput.value = `${parentCategory.id} - ${parentCategory.parentName}`;
        parentInput.dataset.selectedValue = parentCategory.id;
      }
    } else {
      parentInput.value = '';
      parentInput.dataset.selectedValue = '';
    }
  }

  async populateEditDropdowns() {
    if (!this.data || this.data.length === 0) {
      alert('분석된 데이터가 없습니다. 먼저 엑셀 파일을 분석해주세요.');
      return;
    }

    // 부모 카테고리번호 드롭다운 채우기 (비동기 처리)
    await this.populateEditParentDropdown();

    // 도메인구분 드롭다운 채우기
    this.populateEditDomainDropdown();
  }

  async populateEditParentDropdown() {
    const parentInput = document.getElementById('editParentId');
    const dropdownList = document.getElementById('editParentDropdownList');

    // 검색 가능한 드롭다운 초기화
    this.setupEditSearchableDropdown(parentInput, dropdownList);

    // 데이터를 메모리에 저장 (검색용)
    this.editParentDropdownData = this.data.map(item => ({
      id: item.id,
      text: `${item.id} - ${item.parentName}`,
      parentName: item.parentName
    }));

    // 최상위 카테고리 옵션 추가 (기본값으로 설정)
    const topLevelOption = document.createElement('div');
    topLevelOption.className = 'dropdown-item top-level-option';
    topLevelOption.dataset.value = '';
    topLevelOption.textContent = '🏠 최상위 카테고리 (기본 분석 목록)';
    dropdownList.appendChild(topLevelOption);

    // 검색 결과 표시 (처음에는 모든 항목)
    this.showEditDropdownItems(this.editParentDropdownData, dropdownList);
  }

  setupEditSearchableDropdown(input, dropdownList) {
    let isDropdownOpen = false;
    let selectedIndex = -1;

    // 입력 필드 클릭 시 드롭다운 토글
    input.addEventListener('click', () => {
      if (!isDropdownOpen) {
        this.showEditDropdown(dropdownList);
        isDropdownOpen = true;
      }
    });

    // 입력 필드에 포커스 시 드롭다운 열기
    input.addEventListener('focus', () => {
      if (!isDropdownOpen) {
        this.showEditDropdown(dropdownList);
        isDropdownOpen = true;
      }
    });

    // 입력 필드에서 검색
    input.addEventListener('input', (e) => {
      const searchTerm = e.target.value.toLowerCase();
      const filteredData = this.editParentDropdownData.filter(item =>
        item.text.toLowerCase().includes(searchTerm)
      );
      this.showEditDropdownItems(filteredData, dropdownList);
    });

    // 키보드 네비게이션
    input.addEventListener('keydown', (e) => {
      const items = dropdownList.querySelectorAll('.dropdown-item');

      switch (e.key) {
        case 'ArrowDown':
          e.preventDefault();
          selectedIndex = Math.min(selectedIndex + 1, items.length - 1);
          this.highlightEditDropdownItem(items, selectedIndex);
          break;
        case 'ArrowUp':
          e.preventDefault();
          selectedIndex = Math.max(selectedIndex - 1, -1);
          this.highlightEditDropdownItem(items, selectedIndex);
          break;
        case 'Enter':
          e.preventDefault();
          if (selectedIndex >= 0 && items[selectedIndex]) {
            this.selectEditDropdownItem(items[selectedIndex], input);
            this.hideEditDropdown(dropdownList);
            isDropdownOpen = false;
          }
          break;
        case 'Escape':
          this.hideEditDropdown(dropdownList);
          isDropdownOpen = false;
          break;
      }
    });

    // 드롭다운 외부 클릭 시 닫기
    document.addEventListener('click', (e) => {
      if (!input.contains(e.target) && !dropdownList.contains(e.target)) {
        this.hideEditDropdown(dropdownList);
        isDropdownOpen = false;
      }
    });
  }

  showEditDropdown(dropdownList) {
    dropdownList.style.display = 'block';
  }

  hideEditDropdown(dropdownList) {
    dropdownList.style.display = 'none';
  }

  showEditDropdownItems(data, dropdownList) {
    // 기존 항목들 제거 (최상위 옵션 제외)
    const existingItems = dropdownList.querySelectorAll('.dropdown-item:not(:first-child)');
    existingItems.forEach(item => item.remove());

    // 새 항목들 추가
    data.forEach(item => {
      const dropdownItem = document.createElement('div');
      dropdownItem.className = 'dropdown-item';
      dropdownItem.dataset.value = item.id;
      dropdownItem.textContent = item.text;

      dropdownItem.addEventListener('click', () => {
        this.selectEditDropdownItem(dropdownItem, document.getElementById('editParentId'));
        this.hideEditDropdown(dropdownList);
      });

      dropdownList.appendChild(dropdownItem);
    });
  }

  highlightEditDropdownItem(items, index) {
    items.forEach((item, i) => {
      if (i === index) {
        item.classList.add('selected');
      } else {
        item.classList.remove('selected');
      }
    });
  }

  selectEditDropdownItem(item, input) {
    const value = item.dataset.value;
    const text = item.textContent;

    input.value = text;
    input.dataset.selectedValue = value;
  }

  populateEditDomainDropdown() {
    const domainTypeSelect = document.getElementById('editDomainType');
    domainTypeSelect.innerHTML = '<option value="">선택하세요</option>';

    // 도메인 타입 수집 (최적화)
    const domainTypes = new Set();
    for (let i = 0; i < this.data.length; i++) {
      const domainType = this.data[i].domainType;
      if (domainType) {
        domainTypes.add(domainType);
      }
    }

    // 도메인 타입 옵션 추가
    domainTypes.forEach(domainType => {
      const option = document.createElement('option');
      option.value = domainType;
      option.textContent = domainType;
      domainTypeSelect.appendChild(option);
    });
  }

  clearForm() {
    this.categoryForm.reset();

    // 검색 가능한 드롭다운 초기화
    const parentInput = document.getElementById('newParentId');
    if (parentInput) {
      // 선택된 카테고리가 있으면 그 카테고리를 부모로 설정
      if (this.selectedItem) {
        parentInput.value = `${this.selectedItem.id} - ${this.selectedItem.parentName}`;
        parentInput.dataset.selectedValue = this.selectedItem.id;
      } else {
        parentInput.value = '';
        parentInput.dataset.selectedValue = '';
      }
    }

    // 드롭다운 리스트 숨기기
    const dropdownList = document.getElementById('parentDropdownList');
    if (dropdownList) {
      dropdownList.style.display = 'none';
    }

    // 자동으로 다음 ID 채우기
    this.autoFillNextId();
  }

  autoFillNextId() {
    if (!this.data || this.data.length === 0) {
      return;
    }

    // 기존 데이터에서 가장 큰 ID 찾기
    let maxId = 0;
    for (let i = 0; i < this.data.length; i++) {
      const currentId = parseInt(this.data[i].id);
      if (!isNaN(currentId) && currentId > maxId) {
        maxId = currentId;
      }
    }

    // 다음 ID 계산 (가장 큰 ID + 1)
    const nextId = maxId + 1;

    // 일련번호 입력란에 자동으로 채우기
    const idInput = document.getElementById('newId');
    if (idInput) {
      idInput.value = nextId.toString();
    }
  }

  async populateDropdowns() {
    if (!this.data || this.data.length === 0) {
      alert('분석된 데이터가 없습니다. 먼저 엑셀 파일을 분석해주세요.');
      return;
    }

    // 부모 카테고리번호 드롭다운 채우기 (비동기 처리)
    await this.populateParentDropdown();

    // 도메인구분 드롭다운 채우기
    this.populateDomainDropdown();
  }

  async populateParentDropdown() {
    const parentInput = document.getElementById('newParentId');
    const dropdownList = document.getElementById('parentDropdownList');

    // 검색 가능한 드롭다운 초기화
    this.setupSearchableDropdown(parentInput, dropdownList);

    // 데이터를 메모리에 저장 (검색용)
    this.parentDropdownData = this.data.map(item => ({
      id: item.id,
      text: `${item.id} - ${item.parentName}`,
      parentName: item.parentName
    }));

    // 최상위 카테고리 옵션 추가 (기본값으로 설정)
    const topLevelOption = document.createElement('div');
    topLevelOption.className = 'dropdown-item top-level-option';
    topLevelOption.dataset.value = '';
    topLevelOption.textContent = '🏠 최상위 카테고리 (기본 분석 목록)';
    dropdownList.appendChild(topLevelOption);

    // 검색 결과 표시 (처음에는 모든 항목)
    this.showDropdownItems(this.parentDropdownData, dropdownList);
  }

  setupSearchableDropdown(input, dropdownList) {
    let isDropdownOpen = false;
    let selectedIndex = -1;

    // 입력 필드 클릭 시 드롭다운 토글
    input.addEventListener('click', () => {
      if (!isDropdownOpen) {
        this.showDropdown(dropdownList);
        isDropdownOpen = true;
      }
    });

    // 입력 필드에 포커스 시 드롭다운 열기
    input.addEventListener('focus', () => {
      if (!isDropdownOpen) {
        this.showDropdown(dropdownList);
        isDropdownOpen = true;
      }
    });

    // 입력 필드에서 검색
    input.addEventListener('input', (e) => {
      const searchTerm = e.target.value.toLowerCase();
      const filteredData = this.parentDropdownData.filter(item =>
        item.text.toLowerCase().includes(searchTerm)
      );
      this.showDropdownItems(filteredData, dropdownList);
    });

    // 키보드 네비게이션
    input.addEventListener('keydown', (e) => {
      const items = dropdownList.querySelectorAll('.dropdown-item');

      switch (e.key) {
        case 'ArrowDown':
          e.preventDefault();
          selectedIndex = Math.min(selectedIndex + 1, items.length - 1);
          this.highlightDropdownItem(items, selectedIndex);
          break;
        case 'ArrowUp':
          e.preventDefault();
          selectedIndex = Math.max(selectedIndex - 1, -1);
          this.highlightDropdownItem(items, selectedIndex);
          break;
        case 'Enter':
          e.preventDefault();
          if (selectedIndex >= 0 && items[selectedIndex]) {
            this.selectDropdownItem(items[selectedIndex], input);
            this.hideDropdown(dropdownList);
            isDropdownOpen = false;
          }
          break;
        case 'Escape':
          this.hideDropdown(dropdownList);
          isDropdownOpen = false;
          break;
      }
    });

    // 드롭다운 외부 클릭 시 닫기
    document.addEventListener('click', (e) => {
      if (!input.contains(e.target) && !dropdownList.contains(e.target)) {
        this.hideDropdown(dropdownList);
        isDropdownOpen = false;
      }
    });
  }

  showDropdown(dropdownList) {
    dropdownList.style.display = 'block';
  }

  hideDropdown(dropdownList) {
    dropdownList.style.display = 'none';
  }

  showDropdownItems(data, dropdownList) {
    // 기존 항목들 제거 (최상위 옵션 제외)
    const existingItems = dropdownList.querySelectorAll('.dropdown-item:not(:first-child)');
    existingItems.forEach(item => item.remove());

    // 새 항목들 추가
    data.forEach(item => {
      const dropdownItem = document.createElement('div');
      dropdownItem.className = 'dropdown-item';
      dropdownItem.dataset.value = item.id;
      dropdownItem.textContent = item.text;

      dropdownItem.addEventListener('click', () => {
        this.selectDropdownItem(dropdownItem, document.getElementById('newParentId'));
        this.hideDropdown(dropdownList);
      });

      dropdownList.appendChild(dropdownItem);
    });
  }

  highlightDropdownItem(items, index) {
    items.forEach((item, i) => {
      if (i === index) {
        item.classList.add('selected');
      } else {
        item.classList.remove('selected');
      }
    });
  }

  selectDropdownItem(item, input) {
    const value = item.dataset.value;
    const text = item.textContent;

    input.value = text;
    input.dataset.selectedValue = value;

    // 부모 카테고리명 자동 채우기
    if (value && value !== '') {
      const selectedParent = this.data.find(item => item.id === value);
      if (selectedParent) {
        const parentNameInput = document.getElementById('newParentName');
        parentNameInput.value = `${selectedParent.parentName} 하위`;
      }
    } else {
      // 최상위 카테고리 선택 시 부모 카테고리명 초기화
      const parentNameInput = document.getElementById('newParentName');
      parentNameInput.value = '';
    }
  }

  populateDomainDropdown() {
    const domainTypeSelect = document.getElementById('newDomainType');
    domainTypeSelect.innerHTML = '<option value="">선택하세요</option>';

    // 도메인 타입 수집 (최적화)
    const domainTypes = new Set();
    for (let i = 0; i < this.data.length; i++) {
      const domainType = this.data[i].domainType;
      if (domainType) {
        domainTypes.add(domainType);
      }
    }

    // 도메인 타입 옵션 추가
    domainTypes.forEach(domainType => {
      const option = document.createElement('option');
      option.value = domainType;
      option.textContent = domainType;
      domainTypeSelect.appendChild(option);
    });
  }

  async addNewCategory(e) {
    if (e) {
      e.preventDefault();
    }

    const formData = new FormData(this.categoryForm);
    const parentInput = document.getElementById('newParentId');

    const newCategory = {
      id: formData.get('id'),
      parentId: parentInput.dataset.selectedValue || '',
      parentName: formData.get('parentName'),
      domainType: formData.get('domainType') || '',
      hasDomain: formData.get('hasDomain') || '',
      domainAddress: formData.get('domainAddress') || '',
      originalIndex: this.data.length // 기존 데이터의 인덱스 구조 유지
    };

    // parentId가 숫자인 경우 문자열로 변환하여 일관성 유지
    if (newCategory.parentId && !isNaN(newCategory.parentId)) {
      newCategory.parentId = newCategory.parentId.toString();
    }

    // 부모 카테고리가 선택되지 않은 경우 선택된 카테고리를 부모로 설정
    if (!newCategory.parentId || newCategory.parentId === '') {
      if (this.selectedItem) {
        newCategory.parentId = this.selectedItem.id;
        console.log('선택된 카테고리를 부모로 설정:', this.selectedItem.id);
      } else {
        alert('부모 카테고리를 선택해주세요.');
        return;
      }
    }

    // 디버깅: 부모 카테고리 확인
    const parentCategory = this.data.find(item => item.id === newCategory.parentId);
    console.log('새 카테고리 정보:', newCategory);
    console.log('부모 카테고리:', parentCategory);

    // 유효성 검사
    if (!newCategory.id || !newCategory.parentName) {
      alert('일련번호와 카테고리명은 필수 입력 항목입니다.');
      return;
    }

    // 중복 ID 검사
    if (this.data.some(item => item.id === newCategory.id)) {
      alert('이미 존재하는 일련번호입니다. 다른 번호를 사용해주세요.');
      return;
    }

    try {
      this.showLoading(true);

      // 새 카테고리를 데이터에 추가
      this.data.push(newCategory);

      // 트리 구조 재구성
      this.treeData = this.buildTreeStructure();

      // 디버깅: 새로 추가된 카테고리 확인
      console.log('새로 추가된 카테고리:', newCategory);
      console.log('전체 데이터:', this.data);
      console.log('트리 구조:', this.treeData);

      // 통계 업데이트
      this.updateStatistics();

      // 트리 다시 렌더링
      this.renderTree();

      // 성공 메시지
      const parentName = this.selectedItem ? this.selectedItem.parentName : '선택된 카테고리';
      alert(`"${parentName}" 하위에 새 카테고리가 성공적으로 추가되었습니다!`);

      // 입력 섹션 숨기고 분석 섹션 표시
      this.hideInputSection();

      // 업데이트된 엑셀 파일 다운로드 옵션 제공
      if (confirm('업데이트된 엑셀 파일을 다운로드하시겠습니까?')) {
        this.downloadUpdatedExcel();
      }

    } catch (error) {
      console.error('카테고리 추가 중 오류 발생:', error);
      alert('카테고리 추가 중 오류가 발생했습니다.');
    } finally {
      this.showLoading(false);
    }
  }

  getCategoryPath(node) {
    const path = [];
    let currentNode = node;

    // 현재 노드부터 루트까지 경로 구성
    while (currentNode) {
      path.unshift(currentNode.parentName);
      if (currentNode.parentId && currentNode.parentId !== '') {
        currentNode = this.data.find(item => item.id === currentNode.parentId);
      } else {
        break;
      }
    }

    return path.join(' > ');
  }

  showSelectedCategoryInfo(node) {
    // 선택된 카테고리 정보를 표시할 요소 생성 또는 업데이트
    let infoElement = document.getElementById('selectedCategoryInfo');
    if (!infoElement) {
      infoElement = document.createElement('div');
      infoElement.id = 'selectedCategoryInfo';
      infoElement.className = 'selected-category-info';

      // 트리 컨테이너 상단에 추가
      const treeContainer = document.querySelector('.tree-container');
      if (treeContainer) {
        treeContainer.insertBefore(infoElement, treeContainer.firstChild);
      }
    }

    // 선택된 카테고리 경로 표시
    const path = this.getCategoryPath(node);
    infoElement.innerHTML = `
      <div class="selected-info">
        <span class="info-label">선택된 카테고리:</span>
        <span class="info-path">${path}</span>
        <span class="info-hint">← 이 카테고리 하위에 새 카테고리가 추가됩니다</span>
      </div>
    `;
  }

  createProgressUI(message = '펼치는 중...') {
    // 기존 진행 상태 UI 제거
    const existingProgress = document.querySelector('.tree-progress');
    if (existingProgress) {
      existingProgress.remove();
    }

    const progressContainer = document.createElement('div');
    progressContainer.className = 'tree-progress';
    progressContainer.innerHTML = `
            <div class="progress-content">
                <div class="progress-spinner"></div>
                <div class="progress-text">${message}</div>
                <div class="progress-bar">
                    <div class="progress-fill"></div>
                </div>
                <div class="progress-stats">
                    <span class="progress-current">0</span> / <span class="progress-total">0</span>
                </div>
            </div>
        `;

    // 트리 컨테이너에 추가
    this.treeView.appendChild(progressContainer);

    return progressContainer;
  }

  updateProgress(progressContainer, progress, current, total) {
    const progressFill = progressContainer.querySelector('.progress-fill');
    const progressCurrent = progressContainer.querySelector('.progress-current');
    const progressTotal = progressContainer.querySelector('.progress-total');

    if (progressFill) {
      progressFill.style.width = `${progress}%`;
    }

    if (progressCurrent) {
      progressCurrent.textContent = current;
    }

    if (progressTotal) {
      progressTotal.textContent = total;
    }
  }

  hideProgressUI(progressContainer) {
    if (progressContainer && progressContainer.parentNode) {
      progressContainer.style.opacity = '0';
      setTimeout(() => {
        progressContainer.remove();
      }, 300);
    }
  }

  downloadUpdatedExcel() {
    try {
      // 헤더 정의
      const headers = ['일련번호', '부모 카테고리번호', '부모 카테고리명', '도메인구분', '도메인주소여부', '도메인주소'];

      // 데이터 준비
      const excelData = [headers];

      this.data.forEach(item => {
        excelData.push([
          item.id,
          item.parentId,
          item.parentName,
          item.domainType,
          item.hasDomain,
          item.domainAddress
        ]);
      });

      // 워크북 생성
      const workbook = XLSX.utils.book_new();
      const worksheet = XLSX.utils.aoa_to_sheet(excelData);

      // 컬럼 너비 자동 조정
      const colWidths = headers.map(header => ({ wch: Math.max(header.length, 15) }));
      worksheet['!cols'] = colWidths;

      // 워크시트를 워크북에 추가
      XLSX.utils.book_append_sheet(workbook, worksheet, '카테고리');

      // 파일 다운로드
      const fileName = `카테고리_업데이트_${new Date().toISOString().slice(0, 10)}.xlsx`;
      XLSX.writeFile(workbook, fileName);

      // 엑셀 다운로드 후 선택된 카테고리 정보 숨기기
      this.clearSelection();

    } catch (error) {
      console.error('엑셀 파일 생성 중 오류 발생:', error);
      alert('엑셀 파일 생성 중 오류가 발생했습니다.');
    }
  }

  async updateCategory(e) {
    if (e) {
      e.preventDefault();
    }

    const formData = new FormData(this.editCategoryForm);
    const parentInput = document.getElementById('editParentId');

    const updatedCategory = {
      id: formData.get('id'),
      parentId: parentInput.dataset.selectedValue || '',
      parentName: formData.get('parentName'),
      domainType: formData.get('domainType') || '',
      hasDomain: formData.get('hasDomain') || '',
      domainAddress: formData.get('domainAddress') || ''
    };

    // 유효성 검사
    if (!updatedCategory.id || !updatedCategory.parentName) {
      alert('일련번호와 카테고리명은 필수 입력 항목입니다.');
      return;
    }

    // 기본 분석 목록은 수정 불가
    if (updatedCategory.parentName === '기본 분석 목록') {
      alert('기본 분석 목록은 수정할 수 없습니다.');
      return;
    }

    try {
      this.showLoading(true);

      // 데이터에서 해당 카테고리 찾기 및 업데이트
      const categoryIndex = this.data.findIndex(item => item.id === updatedCategory.id);
      if (categoryIndex === -1) {
        alert('수정할 카테고리를 찾을 수 없습니다.');
        return;
      }

      // 기존 데이터 업데이트
      this.data[categoryIndex] = {
        ...this.data[categoryIndex],
        ...updatedCategory
      };

      // 트리 구조 재구성
      this.treeData = this.buildTreeStructure();

      // 통계 업데이트
      this.updateStatistics();

      // 트리 다시 렌더링
      this.renderTree();

      // 성공 메시지
      alert(`"${updatedCategory.parentName}" 카테고리가 성공적으로 수정되었습니다!`);

      // 수정 섹션 숨기고 분석 섹션 표시
      this.hideEditSection();

      // 업데이트된 엑셀 파일 다운로드 옵션 제공
      if (confirm('업데이트된 엑셀 파일을 다운로드하시겠습니까?')) {
        this.downloadUpdatedExcel();
      }

    } catch (error) {
      console.error('카테고리 수정 중 오류 발생:', error);
      alert('카테고리 수정 중 오류가 발생했습니다.');
    } finally {
      this.showLoading(false);
    }
  }

  deleteCategory() {
    // 분석이 완료되지 않았으면 경고 메시지 표시
    if (!this.data || this.data.length === 0) {
      alert('먼저 엑셀 파일을 분석해주세요.');
      return;
    }

    // 선택된 카테고리가 없으면 경고
    if (!this.selectedItem) {
      alert('트리에서 삭제할 카테고리를 먼저 선택해주세요.');
      return;
    }

    // 기본 분석 목록은 삭제 불가
    if (this.selectedItem.parentName === '기본 분석 목록') {
      alert('기본 분석 목록은 삭제할 수 없습니다.');
      return;
    }

    // 하위 카테고리가 있는지 확인
    if (this.selectedItem.children && this.selectedItem.children.length > 0) {
      alert('하위 카테고리가 있는 카테고리는 삭제할 수 없습니다. 먼저 하위 카테고리를 삭제해주세요.');
      return;
    }

    // 삭제 확인
    const categoryName = this.selectedItem.parentName;
    if (!confirm(`"${categoryName}" 카테고리를 정말 삭제하시겠습니까?\n\n이 작업은 되돌릴 수 없습니다.`)) {
      return;
    }

    try {
      this.showLoading(true);

      // 데이터에서 해당 카테고리 제거
      const categoryIndex = this.data.findIndex(item => item.id === this.selectedItem.id);
      if (categoryIndex === -1) {
        alert('삭제할 카테고리를 찾을 수 없습니다.');
        return;
      }

      // 카테고리 삭제
      this.data.splice(categoryIndex, 1);

      // 트리 구조 재구성
      this.treeData = this.buildTreeStructure();

      // 통계 업데이트
      this.updateStatistics();

      // 트리 다시 렌더링
      this.renderTree();

      // 선택 해제
      this.clearSelection();

      // 성공 메시지
      alert(`"${categoryName}" 카테고리가 성공적으로 삭제되었습니다!`);

      // 업데이트된 엑셀 파일 다운로드 옵션 제공
      if (confirm('업데이트된 엑셀 파일을 다운로드하시겠습니까?')) {
        this.downloadUpdatedExcel();
      }

    } catch (error) {
      console.error('카테고리 삭제 중 오류 발생:', error);
      alert('카테고리 삭제 중 오류가 발생했습니다.');
    } finally {
      this.showLoading(false);
    }
  }

  showNoResultsMessage(searchTerm) {
    let noResultsElement = document.getElementById('noResultsMessage');
    if (!noResultsElement) {
      noResultsElement = document.createElement('div');
      noResultsElement.id = 'noResultsMessage';
      noResultsElement.className = 'no-results-message';
      this.treeView.appendChild(noResultsElement);
    }

    noResultsElement.innerHTML = `
      <div class="no-results-content">
        <span class="no-results-icon">🔍</span>
        <span class="no-results-text">"${searchTerm}"에 대한 검색 결과가 없습니다.</span>
      </div>
    `;
    noResultsElement.style.display = 'block';
  }

  hideNoResultsMessage() {
    const noResultsElement = document.getElementById('noResultsMessage');
    if (noResultsElement) {
      noResultsElement.style.display = 'none';
    }
  }

  // 다중 입력 ID 재계산
  recalculateMultiInputIds() {
    if (!this.data || this.data.length === 0) return;

    let maxId = 0;
    for (let i = 0; i < this.data.length; i++) {
      const currentId = parseInt(this.data[i].id);
      if (!isNaN(currentId) && currentId > maxId) {
        maxId = currentId;
      }
    }

    const rows = this.multiFormRows.querySelectorAll('.multi-form-row');
    if (rows.length === 0) return;

    rows.forEach((row, index) => {
      const idInput = row.querySelector('.row-id');
      if (idInput) {
        idInput.value = (maxId + index + 1).toString();
      }
    });
  }

  // 다중 선택 처리
  handleMultiSelect(treeItem, node, isChecked) {
    if (isChecked) {
      // 기본 분석 목록은 선택 불가
      if (node.parentName === '기본 분석 목록') {
        treeItem.querySelector('.tree-checkbox').checked = false;
        alert('기본 분석 목록은 선택할 수 없습니다.');
        return;
      }

      this.selectedItems.push(node);
      treeItem.classList.add('multi-selected');
    } else {
      const index = this.selectedItems.findIndex(item => item.id === node.id);
      if (index !== -1) {
        this.selectedItems.splice(index, 1);
      }
      treeItem.classList.remove('multi-selected');
    }

    // 다중 수정/삭제 버튼 상태 업데이트
    this.multiEditBtn.disabled = this.selectedItems.length === 0;
    this.multiDeleteBtn.disabled = this.selectedItems.length === 0;
  }

  // 다중 선택 모드 토글
  toggleMultiSelectMode() {
    this.isMultiSelectMode = !this.isMultiSelectMode;

    if (this.isMultiSelectMode) {
      this.multiSelectModeBtn.textContent = '단일 선택';
      this.multiSelectModeBtn.title = '단일 선택 모드로 전환';
      this.analysisSection.classList.add('multi-select-active');
      this.selectedItems = [];
      this.renderTree(); // 체크박스가 있는 트리로 다시 렌더링
    } else {
      this.multiSelectModeBtn.textContent = '다중 선택';
      this.multiSelectModeBtn.title = '다중 선택 모드로 전환';
      this.analysisSection.classList.remove('multi-select-active');
      this.selectedItems = [];
      this.renderTree(); // 일반 트리로 다시 렌더링
    }
  }

  // 다중 입력 섹션 표시
  showMultiInputSection() {
    if (!this.data || this.data.length === 0) {
      alert('먼저 엑셀 파일을 분석해주세요.');
      return;
    }

    // 선택된 카테고리가 없으면 경고
    if (!this.selectedItem) {
      alert('트리에서 부모 카테고리를 먼저 선택해주세요.\n\n사용 방법:\n1. 트리에서 원하는 카테고리를 클릭하여 선택\n2. "다중 카테고리 추가" 버튼 클릭\n3. 선택한 카테고리 하위에 새 카테고리들을 추가');
      return;
    }

    this.multiInputSection.style.display = 'block';
    this.analysisSection.style.display = 'none';
    this.clearMultiInputForm();
    this.addMultiInputRow(); // 기본적으로 하나의 행 추가

    // 선택된 카테고리를 첫 번째 행의 부모로 자동 설정
    const firstRow = this.multiFormRows.querySelector('.multi-form-row');
    if (firstRow && this.selectedItem) {
      const parentInput = firstRow.querySelector('.row-parent-id');
      const parentNameInput = firstRow.querySelector('.row-parent-name');

      if (parentInput) {
        parentInput.value = `${this.selectedItem.id} - ${this.selectedItem.parentName}`;
        parentInput.dataset.value = this.selectedItem.id;
      }

      if (parentNameInput) {
        parentNameInput.placeholder = '카테고리명을 입력하세요';
      }
    }

    // 선택된 카테고리 정보 표시
    this.showSelectedCategoryInfoForMultiInput();

    this.recalculateMultiInputIds(); // 초기 ID 설정
  }

  // 다중 입력 섹션 숨기기
  hideMultiInputSection() {
    this.multiInputSection.style.display = 'none';
    this.analysisSection.style.display = 'grid';

    // 선택된 카테고리 정보 숨기기
    const infoElement = document.getElementById('selectedCategoryInfoForMultiInput');
    if (infoElement) {
      infoElement.remove();
    }
  }

  // 다중 수정 섹션 표시
  showMultiEditSection() {
    if (!this.data || this.data.length === 0) {
      alert('먼저 엑셀 파일을 분석해주세요.');
      return;
    }

    if (this.selectedItems.length === 0) {
      alert('다중 선택 모드에서 수정할 카테고리들을 선택해주세요.');
      return;
    }

    this.multiEditSection.style.display = 'block';
    this.analysisSection.style.display = 'none';
    this.populateMultiEditForm();
  }

  // 다중 수정 섹션 숨기기
  hideMultiEditSection() {
    this.multiEditSection.style.display = 'none';
    this.analysisSection.style.display = 'grid';
  }

  // 다중 입력 폼 초기화
  clearMultiInputForm() {
    this.multiFormRows.innerHTML = '';
    this.multiInputRowCount = 0;
  }

  // 다중 입력 행 추가
  addMultiInputRow() {
    this.multiInputRowCount++;
    const row = document.createElement('div');
    row.className = 'multi-form-row';
    row.dataset.rowIndex = this.multiInputRowCount;

    // 새 행 추가 후 모든 행의 ID 재계산

    row.innerHTML = `
      <div class="multi-form-cell checkbox-cell">
        <input type="checkbox" class="row-checkbox" checked>
      </div>
      <div class="multi-form-cell">
        <input type="text" class="row-id" placeholder="자동 생성" readonly>
      </div>
      <div class="multi-form-cell">
        <div class="searchable-dropdown">
          <input type="text" class="row-parent-id" placeholder="부모 카테고리 검색" readonly>
          <div class="dropdown-list" style="display: none;">
            <!-- 동적으로 생성되는 드롭다운 항목들 -->
          </div>
        </div>
      </div>
      <div class="multi-form-cell">
        <input type="text" class="row-parent-name" placeholder="카테고리명" required>
      </div>
      <div class="multi-form-cell">
        <select class="row-domain-type">
          <option value="">선택하세요</option>
        </select>
      </div>
      <div class="multi-form-cell">
        <select class="row-has-domain">
          <option value="">선택하세요</option>
          <option value="Y">Y</option>
          <option value="N">N</option>
        </select>
      </div>
      <div class="multi-form-cell">
        <input type="text" class="row-domain-address" placeholder="도메인 주소">
      </div>
    `;

    this.multiFormRows.appendChild(row);

    // 행 번호 표시
    const rowNumber = this.multiFormRows.children.length;
    row.style.setProperty('--row-number', `"${rowNumber}"`);

    this.populateMultiInputDropdowns(row);
    this.autoFillNextMultiInputId(row);

    // 첫 번째 행이고 선택된 카테고리가 있으면 자동 설정
    if (rowNumber === 1 && this.selectedItem) {
      const parentInput = row.querySelector('.row-parent-id');
      const parentNameInput = row.querySelector('.row-parent-name');

      if (parentInput) {
        parentInput.value = `${this.selectedItem.id} - ${this.selectedItem.parentName}`;
        parentInput.dataset.value = this.selectedItem.id;
      }

      if (parentNameInput) {
        parentNameInput.placeholder = '카테고리명을 입력하세요';
      }
    }

    // 선택된 카테고리가 있으면 모든 새 행에 자동 설정
    if (this.selectedItem) {
      const parentInput = row.querySelector('.row-parent-id');
      const parentNameInput = row.querySelector('.row-parent-name');

      if (parentInput) {
        parentInput.value = `${this.selectedItem.id} - ${this.selectedItem.parentName}`;
        parentInput.dataset.value = this.selectedItem.id;
      }

      if (parentNameInput) {
        parentNameInput.placeholder = '카테고리명을 입력하세요';
      }
    }

    // 부모 카테고리 선택 시 카테고리명 자동 채우기
    const parentInput = row.querySelector('.row-parent-id');
    const parentNameInput = row.querySelector('.row-parent-name');

    // 검색 가능한 드롭다운 설정
    this.setupMultiInputSearchableDropdown(row);

    // Enter 키로 다음 필드로 이동하는 기능 추가
    const inputs = row.querySelectorAll('input, select');
    inputs.forEach((input, index) => {
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          if (index < inputs.length - 1) {
            inputs[index + 1].focus();
          } else {
            // 마지막 필드에서 Enter 시 새 행 추가
            this.addMultiInputRow();
            // 새로 추가된 행의 첫 번째 입력 필드에 포커스
            setTimeout(() => {
              const newRow = this.multiFormRows.lastElementChild;
              const firstInput = newRow.querySelector('input, select');
              if (firstInput) {
                firstInput.focus();
              }
            }, 100);
          }
        }
      });
    });

    // 새 행 추가 후 모든 행의 ID 재계산
    this.recalculateMultiInputIds();
  }

  // 선택된 행 삭제
  removeSelectedRows() {
    const selectedRows = this.multiFormRows.querySelectorAll('.row-checkbox:checked');
    if (selectedRows.length === 0) {
      alert('삭제할 행을 선택해주세요.');
      return;
    }

    selectedRows.forEach(checkbox => {
      checkbox.closest('.multi-form-row').remove();
    });

    // ID 재계산 및 행 번호 업데이트
    this.recalculateMultiInputIds();
    this.updateRowNumbers();
  }

  // 모든 행 지우기
  clearAllMultiInputRows() {
    this.multiFormRows.innerHTML = '';
    this.multiInputRowCount = 0;
    this.updateRowNumbers();

    // 선택된 카테고리 정보가 있으면 다시 표시
    if (this.selectedItem) {
      this.showSelectedCategoryInfoForMultiInput();
    }
  }

  // 행 번호 업데이트
  updateRowNumbers() {
    const rows = this.multiFormRows.querySelectorAll('.multi-form-row');
    rows.forEach((row, index) => {
      row.style.setProperty('--row-number', `"${index + 1}"`);
    });
  }

  // 다중 입력 드롭다운 채우기
  populateMultiInputDropdowns(row) {
    // 부모 카테고리 드롭다운 데이터 준비
    const parentData = this.data
      .filter(item => item.parentName !== '기본 분석 목록') // 기본 분석 목록 제외
      .map(item => ({
        id: item.id,
        parentName: item.parentName,
        text: `${item.id} - ${item.parentName}`
      }))
      .sort((a, b) => a.id - b.id);

    // 도메인 타입 드롭다운 채우기
    const domainTypeSelect = row.querySelector('.row-domain-type');
    if (domainTypeSelect) {
      domainTypeSelect.innerHTML = '<option value="">선택하세요</option>';
      const domainTypes = [...new Set(this.data.map(item => item.domainType).filter(Boolean))];
      domainTypes.forEach(type => {
        const option = document.createElement('option');
        option.value = type;
        option.textContent = type;
        domainTypeSelect.appendChild(option);
      });
    }
  }

  // 다음 ID 자동 채우기
  autoFillNextMultiInputId(row) {
    if (!this.data || this.data.length === 0) return;

    let maxId = 0;
    for (let i = 0; i < this.data.length; i++) {
      const currentId = parseInt(this.data[i].id);
      if (!isNaN(currentId) && currentId > maxId) {
        maxId = currentId;
      }
    }

    // 기존 다중 입력 행들의 ID도 고려
    const existingMultiInputIds = Array.from(this.multiFormRows.querySelectorAll('.row-id'))
      .map(input => parseInt(input.value))
      .filter(id => !isNaN(id));

    if (existingMultiInputIds.length > 0) {
      const maxMultiInputId = Math.max(...existingMultiInputIds);
      maxId = Math.max(maxId, maxMultiInputId);
    }

    const nextId = maxId + 1;
    const idInput = row.querySelector('.row-id');
    if (idInput) {
      idInput.value = nextId.toString();
    }
  }

  // 다중 카테고리 추가
  async addMultipleCategories(e) {
    if (e) {
      e.preventDefault();
    }

    const selectedRows = this.multiFormRows.querySelectorAll('.row-checkbox:checked');
    if (selectedRows.length === 0) {
      alert('추가할 카테고리를 선택해주세요.');
      return;
    }

    const newCategories = [];
    let hasError = false;

    selectedRows.forEach(checkbox => {
      const row = checkbox.closest('.multi-form-row');
      const id = row.querySelector('.row-id').value;
      const parentIdInput = row.querySelector('.row-parent-id');
      const parentId = parentIdInput.dataset.value || '';
      const parentName = row.querySelector('.row-parent-name').value;
      const domainType = row.querySelector('.row-domain-type').value;
      const hasDomain = row.querySelector('.row-has-domain').value;
      const domainAddress = row.querySelector('.row-domain-address').value;

      if (!parentName) {
        alert('카테고리명은 필수 입력 항목입니다.');
        hasError = true;
        return;
      }

      // 기존 데이터와 현재 다중 입력 폼의 다른 행들과 중복 검사
      const existingIds = this.data.map(item => item.id);
      const currentFormIds = Array.from(this.multiFormRows.querySelectorAll('.row-id'))
        .map(input => input.value)
        .filter(inputId => inputId !== id); // 현재 행 제외

      if (existingIds.includes(id) || currentFormIds.includes(id)) {
        alert(`일련번호 ${id}는 이미 존재합니다.`);
        hasError = true;
        return;
      }

      newCategories.push({
        id: id,
        parentId: parentId,
        parentName: parentName,
        domainType: domainType,
        hasDomain: hasDomain,
        domainAddress: domainAddress,
        originalIndex: this.data.length + newCategories.length
      });
    });

    if (hasError) return;

    try {
      this.showLoading(true);

      // 새 카테고리들을 데이터에 추가
      this.data.push(...newCategories);

      // 트리 구조 재구성
      this.treeData = this.buildTreeStructure();

      // 통계 업데이트
      this.updateStatistics();

      // 트리 다시 렌더링
      this.renderTree();

      // 성공 메시지
      alert(`${newCategories.length}개의 카테고리가 성공적으로 추가되었습니다!`);

      // 다중 입력 섹션 숨기고 분석 섹션 표시
      this.hideMultiInputSection();

      // 업데이트된 엑셀 파일 다운로드 옵션 제공
      if (confirm('업데이트된 엑셀 파일을 다운로드하시겠습니까?')) {
        this.downloadUpdatedExcel();
      }

    } catch (error) {
      console.error('다중 카테고리 추가 중 오류 발생:', error);
      alert('다중 카테고리 추가 중 오류가 발생했습니다.');
    } finally {
      this.showLoading(false);
    }
  }

  // 다중 수정 폼 채우기
  populateMultiEditForm() {
    this.multiEditRows.innerHTML = '';

    this.selectedItems.forEach(item => {
      const row = document.createElement('div');
      row.className = 'multi-edit-row';
      row.dataset.itemId = item.id;

      row.innerHTML = `
        <div class="multi-edit-cell checkbox-cell">
          <input type="checkbox" class="edit-row-checkbox" checked>
        </div>
        <div class="multi-edit-cell">
          <input type="text" class="edit-row-id" value="${item.id}" readonly>
        </div>
        <div class="multi-edit-cell">
          <select class="edit-row-parent-id">
            <option value="">선택하세요</option>
          </select>
        </div>
        <div class="multi-edit-cell">
          <input type="text" class="edit-row-parent-name" value="${item.parentName}" required>
        </div>
        <div class="multi-edit-cell">
          <select class="edit-row-domain-type">
            <option value="">선택하세요</option>
          </select>
        </div>
        <div class="multi-edit-cell">
          <select class="edit-row-has-domain">
            <option value="">선택하세요</option>
            <option value="Y">Y</option>
            <option value="N">N</option>
          </select>
        </div>
        <div class="multi-edit-cell">
          <input type="text" class="edit-row-domain-address" value="${item.domainAddress || ''}">
        </div>
      `;

      this.multiEditRows.appendChild(row);
      this.populateMultiEditDropdowns(row, item);
    });
  }

  // 다중 수정 드롭다운 채우기
  populateMultiEditDropdowns(row, item) {
    // 부모 카테고리 드롭다운
    const parentSelect = row.querySelector('.edit-row-parent-id');
    parentSelect.innerHTML = '<option value="">선택하세요</option>';

    this.data.forEach(dataItem => {
      const option = document.createElement('option');
      option.value = dataItem.id;
      option.textContent = `${dataItem.id} - ${dataItem.parentName}`;
      if (dataItem.id === item.parentId) {
        option.selected = true;
      }
      parentSelect.appendChild(option);
    });

    // 도메인 타입 드롭다운
    const domainSelect = row.querySelector('.edit-row-domain-type');
    domainSelect.innerHTML = '<option value="">선택하세요</option>';

    const domainTypes = new Set(this.data.map(dataItem => dataItem.domainType).filter(type => type));
    domainTypes.forEach(domainType => {
      const option = document.createElement('option');
      option.value = domainType;
      option.textContent = domainType;
      if (domainType === item.domainType) {
        option.selected = true;
      }
      domainSelect.appendChild(option);
    });

    // 도메인 주소 여부 드롭다운
    const hasDomainSelect = row.querySelector('.edit-row-has-domain');
    if (item.hasDomain) {
      hasDomainSelect.value = item.hasDomain;
    }
  }

  // 전체 선택
  selectAllEditRows() {
    this.multiEditRows.querySelectorAll('.edit-row-checkbox').forEach(checkbox => {
      checkbox.checked = true;
    });
  }

  // 전체 해제
  deselectAllEditRows() {
    this.multiEditRows.querySelectorAll('.edit-row-checkbox').forEach(checkbox => {
      checkbox.checked = false;
    });
  }

  // 다중 카테고리 수정
  async updateMultipleCategories(e) {
    if (e) {
      e.preventDefault();
    }

    const selectedRows = this.multiEditRows.querySelectorAll('.edit-row-checkbox:checked');
    if (selectedRows.length === 0) {
      alert('수정할 카테고리를 선택해주세요.');
      return;
    }

    const updatedCategories = [];
    let hasError = false;

    selectedRows.forEach(checkbox => {
      const row = checkbox.closest('.multi-edit-row');
      const id = row.querySelector('.edit-row-id').value;
      const parentId = row.querySelector('.edit-row-parent-id').value;
      const parentName = row.querySelector('.edit-row-parent-name').value;
      const domainType = row.querySelector('.edit-row-domain-type').value;
      const hasDomain = row.querySelector('.edit-row-has-domain').value;
      const domainAddress = row.querySelector('.edit-row-domain-address').value;

      if (!parentName) {
        alert('카테고리명은 필수 입력 항목입니다.');
        hasError = true;
        return;
      }

      if (parentName === '기본 분석 목록') {
        alert('기본 분석 목록은 수정할 수 없습니다.');
        hasError = true;
        return;
      }

      updatedCategories.push({
        id: id,
        parentId: parentId,
        parentName: parentName,
        domainType: domainType,
        hasDomain: hasDomain,
        domainAddress: domainAddress
      });
    });

    if (hasError) return;

    try {
      this.showLoading(true);

      // 데이터에서 해당 카테고리들 찾기 및 업데이트
      updatedCategories.forEach(updatedCategory => {
        const categoryIndex = this.data.findIndex(item => item.id === updatedCategory.id);
        if (categoryIndex !== -1) {
          this.data[categoryIndex] = {
            ...this.data[categoryIndex],
            ...updatedCategory
          };
        }
      });

      // 트리 구조 재구성
      this.treeData = this.buildTreeStructure();

      // 통계 업데이트
      this.updateStatistics();

      // 트리 다시 렌더링
      this.renderTree();

      // 성공 메시지
      alert(`${updatedCategories.length}개의 카테고리가 성공적으로 수정되었습니다!`);

      // 다중 수정 섹션 숨기고 분석 섹션 표시
      this.hideMultiEditSection();

      // 업데이트된 엑셀 파일 다운로드 옵션 제공
      if (confirm('업데이트된 엑셀 파일을 다운로드하시겠습니까?')) {
        this.downloadUpdatedExcel();
      }

    } catch (error) {
      console.error('다중 카테고리 수정 중 오류 발생:', error);
      alert('다중 카테고리 수정 중 오류가 발생했습니다.');
    } finally {
      this.showLoading(false);
    }
  }

  // 다중 카테고리 삭제
  multiDeleteCategories() {
    if (!this.data || this.data.length === 0) {
      alert('먼저 엑셀 파일을 분석해주세요.');
      return;
    }

    if (this.selectedItems.length === 0) {
      alert('다중 선택 모드에서 삭제할 카테고리들을 선택해주세요.');
      return;
    }

    // 기본 분석 목록 삭제 방지
    const hasBasicRoot = this.selectedItems.some(item => item.parentName === '기본 분석 목록');
    if (hasBasicRoot) {
      alert('기본 분석 목록은 삭제할 수 없습니다.');
      return;
    }

    // 하위 카테고리가 있는 항목 확인
    const hasChildren = this.selectedItems.some(item => item.children && item.children.length > 0);
    if (hasChildren) {
      alert('하위 카테고리가 있는 카테고리는 삭제할 수 없습니다. 먼저 하위 카테고리를 삭제해주세요.');
      return;
    }

    const categoryNames = this.selectedItems.map(item => item.parentName).join(', ');
    if (!confirm(`다음 카테고리들을 정말 삭제하시겠습니까?\n\n${categoryNames}\n\n이 작업은 되돌릴 수 없습니다.`)) {
      return;
    }

    try {
      this.showLoading(true);

      // 선택된 카테고리들 삭제
      this.selectedItems.forEach(selectedItem => {
        const categoryIndex = this.data.findIndex(item => item.id === selectedItem.id);
        if (categoryIndex !== -1) {
          this.data.splice(categoryIndex, 1);
        }
      });

      // 트리 구조 재구성
      this.treeData = this.buildTreeStructure();

      // 통계 업데이트
      this.updateStatistics();

      // 트리 다시 렌더링
      this.renderTree();

      // 선택 해제
      this.selectedItems = [];
      this.toggleMultiSelectMode(); // 다중 선택 모드 해제

      // 성공 메시지
      const deletedCount = this.selectedItems.length;
      alert(`${deletedCount}개의 카테고리가 성공적으로 삭제되었습니다!`);

      // 업데이트된 엑셀 파일 다운로드 옵션 제공
      if (confirm('업데이트된 엑셀 파일을 다운로드하시겠습니까?')) {
        this.downloadUpdatedExcel();
      }

    } catch (error) {
      console.error('다중 카테고리 삭제 중 오류 발생:', error);
      alert('다중 카테고리 삭제 중 오류가 발생했습니다.');
    } finally {
      this.showLoading(false);
    }
  }

  // 다중 입력용 검색 가능한 드롭다운 설정
  setupMultiInputSearchableDropdown(row) {
    const parentInput = row.querySelector('.row-parent-id');
    const dropdownList = row.querySelector('.dropdown-list');

    if (!parentInput || !dropdownList) return;

    // 부모 카테고리 드롭다운 데이터 준비 (기본 분석 목록 제외)
    const parentData = this.data
      .filter(item => item.parentName !== '기본 분석 목록')
      .map(item => ({
        id: item.id,
        parentName: item.parentName,
        text: `${item.id} - ${item.parentName}`
      }))
      .sort((a, b) => a.id - b.id);

    // 클릭 이벤트 설정
    parentInput.addEventListener('click', () => {
      this.showMultiInputDropdown(dropdownList);
      this.showMultiInputDropdownItems(parentData, dropdownList);
    });

    // 포커스 이벤트 설정
    parentInput.addEventListener('focus', () => {
      this.showMultiInputDropdown(dropdownList);
      this.showMultiInputDropdownItems(parentData, dropdownList);
    });

    // 검색 기능
    parentInput.addEventListener('input', (e) => {
      const searchTerm = e.target.value.toLowerCase();
      const filteredData = parentData.filter(item =>
        item.text.toLowerCase().includes(searchTerm)
      );
      this.showMultiInputDropdownItems(filteredData, dropdownList);
    });

    // 키보드 네비게이션
    let highlightedIndex = -1;
    parentInput.addEventListener('keydown', (e) => {
      const items = dropdownList.querySelectorAll('.dropdown-item');

      switch (e.key) {
        case 'ArrowDown':
          e.preventDefault();
          highlightedIndex = Math.min(highlightedIndex + 1, items.length - 1);
          this.highlightMultiInputDropdownItem(items, highlightedIndex);
          break;
        case 'ArrowUp':
          e.preventDefault();
          highlightedIndex = Math.max(highlightedIndex - 1, -1);
          this.highlightMultiInputDropdownItem(items, highlightedIndex);
          break;
        case 'Enter':
          e.preventDefault();
          if (highlightedIndex >= 0 && items[highlightedIndex]) {
            const parentNameInput = row.querySelector('.row-parent-name');
            this.selectMultiInputDropdownItem(items[highlightedIndex], parentInput, parentNameInput);
            this.hideMultiInputDropdown(dropdownList);
          }
          break;
        case 'Escape':
          this.hideMultiInputDropdown(dropdownList);
          break;
      }
    });

    // 외부 클릭 시 드롭다운 숨기기
    document.addEventListener('click', (e) => {
      if (!parentInput.contains(e.target) && !dropdownList.contains(e.target)) {
        this.hideMultiInputDropdown(dropdownList);
      }
    });
  }

  // 다중 입력 드롭다운 표시/숨김 함수들
  showMultiInputDropdown(dropdownList) {
    dropdownList.style.display = 'block';
  }

  hideMultiInputDropdown(dropdownList) {
    dropdownList.style.display = 'none';
  }

  showMultiInputDropdownItems(data, dropdownList) {
    dropdownList.innerHTML = '';

    // 최상위 카테고리 옵션 추가 (다중 입력에서는 제외)
    // const topLevelOption = document.createElement('div');
    // topLevelOption.className = 'dropdown-item top-level-option';
    // topLevelOption.dataset.value = '';
    // topLevelOption.textContent = '🏠 최상위 카테고리 (기본 분석 목록)';
    // dropdownList.appendChild(topLevelOption);

    // 검색 결과 표시
    data.forEach(item => {
      // 기본 분석 목록은 제외
      if (item.parentName === '기본 분석 목록') {
        return;
      }

      const option = document.createElement('div');
      option.className = 'dropdown-item';
      option.dataset.value = item.id;
      option.dataset.parentName = item.parentName;
      option.textContent = item.text;

      option.addEventListener('click', () => {
        const parentInput = dropdownList.previousElementSibling;
        const parentNameInput = dropdownList.parentElement.parentElement.nextElementSibling.querySelector('.row-parent-name');
        this.selectMultiInputDropdownItem(option, parentInput, parentNameInput);
        this.hideMultiInputDropdown(dropdownList);
      });

      dropdownList.appendChild(option);
    });
  }

  highlightMultiInputDropdownItem(items, index) {
    items.forEach(item => item.classList.remove('highlighted'));
    if (index >= 0 && items[index]) {
      items[index].classList.add('highlighted');
    }
  }

  selectMultiInputDropdownItem(item, input, parentNameInput) {
    const selectedValue = item.dataset.value;
    const selectedText = item.textContent;

    input.value = selectedText;
    input.dataset.value = selectedValue;

    if (selectedValue) {
      const selectedParent = this.data.find(dataItem => dataItem.id === selectedValue);
      if (selectedParent) {
        parentNameInput.value = `${selectedParent.parentName} 하위`;
      }
    } else {
      parentNameInput.value = '';
    }
  }

  showSelectedCategoryInfoForMultiInput() {
    // 선택된 카테고리 정보를 표시할 요소 생성 또는 업데이트
    let infoElement = document.getElementById('selectedCategoryInfoForMultiInput');
    if (!infoElement) {
      infoElement = document.createElement('div');
      infoElement.id = 'selectedCategoryInfoForMultiInput';
      infoElement.className = 'selected-category-info';

      // 다중 입력 섹션 상단에 추가
      const multiInputContainer = document.querySelector('.multi-input-container');
      if (multiInputContainer) {
        const h3 = multiInputContainer.querySelector('h3');
        if (h3) {
          multiInputContainer.insertBefore(infoElement, h3.nextSibling);
        }
      }
    }

    // 선택된 카테고리 경로 표시
    const path = this.getCategoryPath(this.selectedItem);
    infoElement.innerHTML = `
      <div class="selected-info">
        <span class="info-label">📁 선택된 부모 카테고리:</span>
        <span class="info-path">${path}</span>
        <span class="info-hint">← 새로 추가되는 모든 행의 부모 카테고리로 자동 설정됩니다</span>
      </div>
    `;

    // 정보 요소를 보이게 함
    infoElement.style.display = 'block';
  }

  // 분석 통계 섹션 펼쳐보기/접기 기능
  toggleStatsSection() {
    if (!this.statsContent || !this.statsToggleBtn) {
      return;
    }
    
    const isCollapsed = this.statsContent.classList.contains('collapsed');
    
    if (isCollapsed) {
      // 펼치기
      this.statsContent.classList.remove('collapsed');
      this.statsToggleBtn.classList.remove('collapsed');
      this.statsContent.style.display = 'block';
      this.statsContent.style.opacity = '1';
      this.statsContent.style.transform = 'translateY(0)';
    } else {
      // 접기
      this.statsContent.style.opacity = '0';
      this.statsContent.style.transform = 'translateY(-10px)';
      this.statsContent.classList.add('collapsed');
      this.statsToggleBtn.classList.add('collapsed');
    }
  }
}

// 애플리케이션 초기화
document.addEventListener('DOMContentLoaded', () => {
  new CategoryAnalyzer();
});