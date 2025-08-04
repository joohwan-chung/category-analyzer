class CategoryDownloader {
  constructor() {
    this.analyzer = null; // CategoryAnalyzer 인스턴스 참조
    this.initializeElements();
    this.bindEvents();
  }

  // CategoryAnalyzer 인스턴스 설정
  setAnalyzer(analyzer) {
    this.analyzer = analyzer;
  }

  initializeElements() {
    this.downloadBtn = document.getElementById('downloadBtn');
  }

  bindEvents() {
    if (this.downloadBtn) {
      this.downloadBtn.addEventListener('click', this.downloadCategories.bind(this));
    }
  }

  // 분석 완료 후 다운로드 버튼 활성화
  enableDownloadButton() {
    if (this.downloadBtn) {
      this.downloadBtn.disabled = false;
      this.downloadBtn.title = "클릭하여 카테고리 엑셀 파일을 다운로드하세요";
    }
  }

  // 분석 시작 시 다운로드 버튼 비활성화
  disableDownloadButton() {
    if (this.downloadBtn) {
      this.downloadBtn.disabled = true;
      this.downloadBtn.title = "먼저 엑셀 파일을 분석한 후 사용할 수 있습니다";
    }
  }

  // 카테고리 경로 생성 (부모 카테고리들을 모두 포함)
  getCategoryPath(item, data) {
    const path = [];
    let currentItem = item;

    // 현재 아이템부터 루트까지 경로 구성
    while (currentItem) {
      path.unshift(currentItem.parentName);
      if (currentItem.parentId && currentItem.parentId !== '' && currentItem.parentId !== '0') {
        currentItem = data.find(dataItem => dataItem.id === currentItem.parentId);
      } else {
        break;
      }
    }

    return path;
  }

  // 카테고리 다운로드 처리 (요청하신 구조에 맞춤)
  async downloadCategories() {
    if (!this.analyzer || !this.analyzer.data || this.analyzer.data.length === 0) {
      alert('먼저 엑셀 파일을 분석해주세요.');
      return;
    }

    try {
      // 프로그레스 바 UI 생성
      const progressContainer = this.createDownloadProgressUI();
      
      // 1단계: 최대 깊이 계산 (10%)
      this.updateDownloadProgress(progressContainer, 10, '최대 깊이 계산 중...');
      await this.delay(100);
      
      let maxDepth = 0;
      this.analyzer.data.forEach(item => {
        const path = this.getCategoryPath(item, this.analyzer.data);
        maxDepth = Math.max(maxDepth, path.length);
      });

      // 2단계: 헤더 정의 (20%)
      this.updateDownloadProgress(progressContainer, 20, '헤더 생성 중...');
      await this.delay(100);
      
      const headers = ['일련번호', '부모 카테고리번호'];
      for (let i = 1; i <= maxDepth; i++) {
        headers.push(`레벨${i} 카테고리명`);
      }
      headers.push('도메인주소');

      // 3단계: 데이터 준비 (60%)
      this.updateDownloadProgress(progressContainer, 30, '데이터 처리 중...');
      await this.delay(100);
      
      const excelData = [headers];
      const totalItems = this.analyzer.data.length;
      
      // 배치 처리로 진행률 업데이트
      const batchSize = Math.max(1, Math.floor(totalItems / 20)); // 20단계로 나누기
      let processedItems = 0;
      
      for (let i = 0; i < this.analyzer.data.length; i++) {
        const item = this.analyzer.data[i];
        const categoryPath = this.getCategoryPath(item, this.analyzer.data);
        
        // 기본 데이터
        const row = [item.id, item.parentId];
        
        // 각 레벨별 카테고리명 추가
        for (let j = 0; j < maxDepth; j++) {
          row.push(categoryPath[j] || ''); // 해당 레벨에 카테고리가 없으면 빈 문자열
        }
        
        // 도메인 주소를 맨 끝에 추가
        row.push(item.domainAddress || '');
        
        excelData.push(row);
        processedItems++;
        
        // 배치 단위로 진행률 업데이트
        if (processedItems % batchSize === 0 || processedItems === totalItems) {
          const progress = 30 + (processedItems / totalItems) * 30; // 30% ~ 60%
          this.updateDownloadProgress(progressContainer, progress, `데이터 처리 중... (${processedItems}/${totalItems})`);
          await this.delay(50);
        }
      }

      // 4단계: 워크북 생성 (80%)
      this.updateDownloadProgress(progressContainer, 80, '엑셀 파일 생성 중...');
      await this.delay(100);
      
      const workbook = XLSX.utils.book_new();
      const worksheet = XLSX.utils.aoa_to_sheet(excelData);

      // 컬럼 너비 자동 조정
      const colWidths = headers.map(header => ({ wch: Math.max(header.length, 15) }));
      worksheet['!cols'] = colWidths;

      // 워크시트를 워크북에 추가
      XLSX.utils.book_append_sheet(workbook, worksheet, '카테고리');

      // 5단계: 파일 다운로드 (90%)
      this.updateDownloadProgress(progressContainer, 90, '파일 다운로드 준비 중...');
      await this.delay(100);
      
      const fileName = `카테고리_분석_${new Date().toISOString().slice(0, 10)}.xlsx`;
      
      // 6단계: 다운로드 완료 (100%)
      this.updateDownloadProgress(progressContainer, 100, '다운로드 완료!');
      await this.delay(500);
      
      XLSX.writeFile(workbook, fileName);

      // 프로그레스 바 숨기기
      this.hideDownloadProgressUI(progressContainer);

      // 성공 메시지
      alert('카테고리 엑셀 파일이 성공적으로 다운로드되었습니다!\n\n다운로드된 파일에는 각 카테고리의 레벨별 경로가 별도 열로 표시됩니다.');

    } catch (error) {
      console.error('엑셀 파일 생성 중 오류 발생:', error);
      alert('엑셀 파일 생성 중 오류가 발생했습니다.');
    }
  }

  // 다운로드 프로그레스 바 UI 생성
  createDownloadProgressUI() {
    // 기존 프로그레스 바 제거
    const existingProgress = document.querySelector('.download-progress');
    if (existingProgress) {
      existingProgress.remove();
    }

    const progressContainer = document.createElement('div');
    progressContainer.className = 'download-progress';
    progressContainer.innerHTML = `
      <div class="download-progress-content">
        <div class="download-progress-header">
          <div class="download-progress-icon">📊</div>
          <div class="download-progress-title">카테고리 엑셀 파일 다운로드</div>
        </div>
        <div class="download-progress-bar">
          <div class="download-progress-fill"></div>
        </div>
        <div class="download-progress-text">준비 중...</div>
        <div class="download-progress-percentage">0%</div>
      </div>
    `;

    // 다운로드 버튼 근처에 추가
    if (this.downloadBtn) {
      this.downloadBtn.parentNode.insertBefore(progressContainer, this.downloadBtn.nextSibling);
    } else {
      document.body.appendChild(progressContainer);
    }

    return progressContainer;
  }

  // 다운로드 진행률 업데이트
  updateDownloadProgress(progressContainer, progress, text) {
    const progressFill = progressContainer.querySelector('.download-progress-fill');
    const progressText = progressContainer.querySelector('.download-progress-text');
    const progressPercentage = progressContainer.querySelector('.download-progress-percentage');

    if (progressFill) {
      progressFill.style.width = `${progress}%`;
    }

    if (progressText) {
      progressText.textContent = text;
    }

    if (progressPercentage) {
      progressPercentage.textContent = `${Math.round(progress)}%`;
    }
  }

  // 다운로드 프로그레스 바 숨기기
  hideDownloadProgressUI(progressContainer) {
    if (progressContainer) {
      progressContainer.style.opacity = '0';
      setTimeout(() => {
        if (progressContainer.parentNode) {
          progressContainer.remove();
        }
      }, 500);
    }
  }

  // 지연 함수 (UI 업데이트를 위한 비동기 처리)
  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }


}

// 전역 인스턴스 생성
window.categoryDownloader = new CategoryDownloader(); 