// 发票整理工具主应用
class InvoiceManager {
  constructor() {
    // 初始化变量
    this.invoices = [];
    this.currentPageIndex = 0; // 当前预览页面索引
    this.previewPages = []; // 预览页面数组
    
    this.paperConfig = {
      type: 'A4',
      orientation: 'portrait'
    };
    this.layoutConfig = {
      invoicesPerSheet: 3
    };
    
    // 命名配置
    this.namingConfig = {
      projectName: '',
      product: '',
      projectCycleStart: '',
      projectCycleEnd: ''
    };
    
    // 初始化应用
    this.init();
  }

  init() {
    // 获取DOM元素
    this.dropZone = document.getElementById('dropZone');
    this.fileInput = document.getElementById('fileInput');
    this.browseBtn = document.getElementById('browseBtn');
    this.paperType = document.getElementById('paperType');
    this.orientation = document.getElementById('orientation');
    this.customSize = document.getElementById('customSize');
    this.customWidth = document.getElementById('customWidth');
    this.customHeight = document.getElementById('customHeight');
    this.invoicePerSheet = document.getElementById('invoicePerSheet');
    this.invoicesList = document.getElementById('invoicesList');
    this.invoiceCount = document.getElementById('invoiceCount');
    this.previewBtn = document.getElementById('previewBtn');
    this.exportBtn = document.getElementById('exportBtn');
    this.previewCanvas = document.getElementById('previewCanvas');
    this.loadingOverlay = document.getElementById('loadingOverlay');
    
    // 全屏预览相关元素
    this.fullscreenPreview = document.getElementById('fullscreenPreview');
    this.closePreviewBtn = document.getElementById('closePreviewBtn');
    this.prevPageBtn = document.getElementById('prevPageBtn');
    this.nextPageBtn = document.getElementById('nextPageBtn');
    this.currentPageSpan = document.getElementById('currentPage');
    this.totalPagesSpan = document.getElementById('totalPages');

    // 新增命名设置元素
    this.projectNameInput = document.getElementById('projectName');
    this.productInput = document.getElementById('product');
    this.projectCycleStartInput = document.getElementById('projectCycleStart');
    this.projectCycleEndInput = document.getElementById('projectCycleEnd');

    // 检查关键元素是否存在
    if (!this.previewCanvas || !this.currentPageSpan || !this.totalPagesSpan) {
      console.error('关键预览元素未找到，请检查HTML结构');
    }
    
    // 绑定事件
    this.bindEvents();
    
    // 初始化PDF.js
    this.initializePDFJS();
    
    // 初始化拖拽功能
    this.initDragAndDrop();
  }

  initializePDFJS() {
    // 检查PDF.js是否已加载
    if (typeof pdfjsLib !== 'undefined') {
      // 设置worker源
      pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.4.120/build/pdf.worker.min.js';
      console.log('PDF.js已初始化');
    } else {
      console.error('PDF.js库未加载');
      // 尝试延迟加载
      setTimeout(() => {
        if (typeof pdfjsLib !== 'undefined') {
          pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.4.120/build/pdf.worker.min.js';
          console.log('PDF.js已通过延迟加载初始化');
        } else {
          console.error('PDF.js库始终未能加载，请检查网络连接');
        }
      }, 2000);
    }
  }

  bindEvents() {
    // 文件上传相关事件
    this.browseBtn.addEventListener('click', (e) => {
      e.preventDefault(); // 阻止默认行为
      this.fileInput.click();
    });
    this.fileInput.addEventListener('change', (e) => this.handleFiles(e.target.files));
    this.dropZone.addEventListener('dragover', (e) => {
      e.preventDefault();
      this.dropZone.classList.add('dragover');
    });
    this.dropZone.addEventListener('dragenter', (e) => {
      e.preventDefault();
      this.dropZone.classList.add('dragover');
    });
    this.dropZone.addEventListener('dragleave', () => {
      this.dropZone.classList.remove('dragover');
    });
    this.dropZone.addEventListener('drop', (e) => {
      e.preventDefault();
      this.dropZone.classList.remove('dragover');
      // 添加动画效果
      this.dropZone.classList.add('animate');
      setTimeout(() => {
        this.dropZone.classList.remove('animate');
      }, 600);
      this.handleFiles(e.dataTransfer.files);
    });

    // 设置变更事件
    this.paperType.addEventListener('change', (e) => {
      this.paperConfig.type = e.target.value;
    });
    
    this.orientation.addEventListener('change', (e) => {
      this.paperConfig.orientation = e.target.value;
    });
    
    this.invoicePerSheet.addEventListener('change', (e) => {
      this.layoutConfig.invoicesPerSheet = parseInt(e.target.value);
    });
    
    // 命名设置事件
    this.projectNameInput.addEventListener('input', (e) => {
      this.namingConfig.projectName = e.target.value;
    });
    
    this.productInput.addEventListener('input', (e) => {
      this.namingConfig.product = e.target.value;
    });
    
    this.projectCycleStartInput.addEventListener('input', (e) => {
      this.namingConfig.projectCycleStart = e.target.value;
    });
    
    this.projectCycleEndInput.addEventListener('input', (e) => {
      this.namingConfig.projectCycleEnd = e.target.value;
    });
    
    // 按钮事件
    this.previewBtn.addEventListener('click', () => this.generatePreview());
    this.exportBtn.addEventListener('click', () => this.exportPDF());
    
    // 全屏预览关闭事件 - 确保元素存在
    if (this.closePreviewBtn) {
      this.closePreviewBtn.addEventListener('click', () => {
        this.fullscreenPreview.style.display = 'none';
      });
    } else {
      console.error('关闭预览按钮未找到');
    }
    
    // 翻页按钮事件 - 确保元素存在
    if (this.prevPageBtn) {
      this.prevPageBtn.addEventListener('click', () => this.previousPage());
    } else {
      console.error('上一页按钮未找到');
    }
    
    if (this.nextPageBtn) {
      this.nextPageBtn.addEventListener('click', () => this.nextPage());
    } else {
      console.error('下一页按钮未找到');
    }
    
    // 点击模态框外部关闭预览 - 确保元素存在
    if (this.fullscreenPreview) {
      this.fullscreenPreview.addEventListener('click', (e) => {
        if (e.target === this.fullscreenPreview) {
          this.fullscreenPreview.style.display = 'none';
        }
      });
    } else {
      console.error('全屏预览容器未找到');
    }
  }

  // 计算文件哈希值的辅助方法
  async calculateFileHash(file) {
    // 检查浏览器是否支持crypto API
    if (!crypto || !crypto.subtle || !crypto.subtle.digest) {
      // 如果不支持crypto API，使用一个简单的替代方法
      const buffer = await file.arrayBuffer();
      let hash = 0;
      const bytes = new Uint8Array(buffer);
      for (let i = 0; i < bytes.length; i++) {
        hash = ((hash << 5) - hash) + bytes[i];
        hash |= 0; // 转换为32位整数
      }
      return hash.toString();
    }
    
    const buffer = await file.arrayBuffer();
    const hashBuffer = await crypto.subtle.digest('SHA-256', buffer);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
  }

  async handleFiles(files) {
    this.showLoading(true);
    
    // 收集重复文件名以便一次性提示
    const duplicateFiles = [];
    
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      
      try {
        // 检查是否已存在相同文件
        const fileHash = await this.calculateFileHash(file);
        const isDuplicate = this.invoices.some(invoice => invoice.hash === fileHash);
        
        if (isDuplicate) {
          duplicateFiles.push(file.name);
          continue;
        }
        
        let invoiceData = null;
        
        // 改进文件类型检测
        if (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')) {
          invoiceData = await this.processPDF(file);
        } else if (file.type.includes('image') || file.name.toLowerCase().endsWith('.jpg') || 
                   file.name.toLowerCase().endsWith('.jpeg') || file.name.toLowerCase().endsWith('.png') ||
                   file.name.toLowerCase().endsWith('.gif') || file.name.toLowerCase().endsWith('.bmp')) {
          invoiceData = await this.processImage(file);
        } else if (file.name.toLowerCase().endsWith('.ofd')) {
          // OFD文件处理（简化模拟）
          invoiceData = await this.processOFD(file);
        } else {
          console.error(`不支持的文件类型: ${file.type}`);
          alert(`不支持的文件类型: ${file.name} (${file.type})`);
          continue;
        }
        
        if (invoiceData) {
          // 添加文件哈希值
          invoiceData.hash = fileHash;
          this.invoices.push(invoiceData);
          this.renderInvoiceItem(invoiceData);
          
          // 添加成功提示
          console.log(`文件 "${file.name}" 上传成功`);
        }
      } catch (error) {
        console.error(`处理文件 ${file.name} 时出错:`, error);
        alert(`处理文件 ${file.name} 时出错: ${error.message}`);
      }
    }
    
    // 如果有重复文件，一次性提示用户
    if (duplicateFiles.length > 0) {
      const message = `以下 ${duplicateFiles.length} 个文件已存在，不能重复上传：\n\n${duplicateFiles.join('\n')}\n\n共${files.length}个文件，成功上传${files.length - duplicateFiles.length}个。`;
      alert(message);
    } else if (files.length > 0) {
      // 显示成功上传的数量
      alert(`成功上传 ${files.length - duplicateFiles.length} 个文件`);
    }
    
    this.updateInvoiceCount();
    this.showLoading(false);
  }

  async processPDF(file) {
    let url = null;
    try {
      // 确保PDF.js已加载
      if (typeof pdfjsLib === 'undefined') {
        throw new Error('PDF.js库未加载，请检查网络连接');
      }
      
      // 创建临时URL
      url = URL.createObjectURL(file);
      
      // 使用更全面的配置选项处理各种类型的PDF
      const loadingTask = pdfjsLib.getDocument({
        url: url,
        cMapUrl: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.4.120/cmaps/',
        cMapPacked: true,
        enableXfa: true, // 启用XFA表单支持
        verbosity: pdfjsLib.VerbosityLevel.ERROR // 只显示错误级别日志
      });
      
      const pdf = await loadingTask.promise;
      const page = await pdf.getPage(1);
      
      // 获取页面尺寸，使用更高的分辨率以获得更好的质量
      // 提高缩放比例以获得更清晰的图像
      const scale = 3.0; // 从1.5提高到3.0以显著提升清晰度
      const viewport = page.getViewport({ scale: scale });
      
      // 创建canvas用于渲染PDF页面
      const canvas = document.createElement('canvas');
      const context = canvas.getContext('2d');
      canvas.height = viewport.height;
      canvas.width = viewport.width;
      
      const renderContext = {
        canvasContext: context,
        viewport: viewport
      };
      
      // 渲染页面到canvas
      await page.render(renderContext).promise;
      
      // 转换为数据URL - 使用PNG格式以保留所有细节，质量设置为最高
      const dataUrl = canvas.toDataURL('image/png', 1.0); // 使用PNG格式和最高质量
      
      return {
        id: Date.now() + Math.random(),
        name: file.name,
        type: 'pdf',
        size: file.size,
        dataUrl: dataUrl,
        thumbnail: dataUrl, // 使用相同图像作为缩略图
        metadata: {}
      };
    } catch (error) {
      console.error('处理PDF失败:', error);
      
      // 确保释放URL对象
      if (url) {
        try {
          URL.revokeObjectURL(url);
        } catch (revokeError) {
          console.warn('释放URL对象时出错:', revokeError);
        }
      }
      
      // 尝试使用备用方法
      try {
        // 创建一个表示错误的图像
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        canvas.width = 400;
        canvas.height = 200;
        
        // 绘制错误提示
        ctx.fillStyle = '#f0f0f0';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.fillStyle = '#ff0000';
        ctx.font = '14px Arial';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText('PDF处理失败', canvas.width/2, canvas.height/2 - 10);
        ctx.fillText(file.name, canvas.width/2, canvas.height/2 + 10);
        
        const dataUrl = canvas.toDataURL('image/png');
        
        return {
          id: Date.now() + Math.random(),
          name: file.name,
          type: 'pdf',
          size: file.size,
          dataUrl: dataUrl,
          thumbnail: dataUrl,
          metadata: {}
        };
      } catch (fallbackError) {
        console.error('PDF处理备用方法也失败:', fallbackError);
        throw error; // 重新抛出原始错误
      }
    } finally {
      // 确保在正常流程中也释放URL对象
      if (url) {
        try {
          URL.revokeObjectURL(url);
        } catch (revokeError) {
          console.warn('释放URL对象时出错:', revokeError);
        }
      }
    }
  }

  async processImage(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      
      reader.onload = (e) => {
        const img = new Image();
        
        img.onload = () => {
          // 创建缩略图
          const thumbCanvas = document.createElement('canvas');
          const thumbCtx = thumbCanvas.getContext('2d');
          const maxWidth = 200;
          const maxHeight = 200;
          
          let { width, height } = this.calculateThumbnailSize(img.width, img.height, maxWidth, maxHeight);
          
          thumbCanvas.width = width;
          thumbCanvas.height = height;
          
          thumbCtx.drawImage(img, 0, 0, width, height);
          const thumbnail = thumbCanvas.toDataURL('image/jpeg');
          
          resolve({
            id: Date.now() + Math.random(),
            name: file.name,
            type: 'image',
            size: file.size,
            dataUrl: e.target.result,
            thumbnail: thumbnail,
            metadata: {}
          });
        };
        
        img.onerror = reject;
        img.src = e.target.result;
      };
      
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  async processOFD(file) {
    // 模拟OFD处理（由于缺少OFD库，暂时使用图像处理逻辑）
    return this.processImage(file);
  }

  calculateThumbnailSize(width, height, maxWidth, maxHeight) {
    if (width <= maxWidth && height <= maxHeight) {
      return { width, height };
    }
    
    const ratio = Math.min(maxWidth / width, maxHeight / height);
    return {
      width: Math.round(width * ratio),
      height: Math.round(height * ratio)
    };
  }

  renderInvoiceItem(invoice) {
    const invoiceEl = document.createElement('div');
    invoiceEl.className = 'invoice-item';
    invoiceEl.setAttribute('data-id', invoice.id); // 添加ID属性用于拖拽
    invoiceEl.draggable = true; // 启用拖拽
    invoiceEl.innerHTML = `
      <div class="info">
        <p><strong>${invoice.name}</strong></p>
        <p>类型: ${invoice.type.toUpperCase()} | ${(invoice.size / 1024).toFixed(2)} KB</p>
      </div>
      <button class="remove-btn" data-id="${invoice.id}">×</button>
    `;
    
    // 添加删除事件
    const removeBtn = invoiceEl.querySelector('.remove-btn');
    removeBtn.addEventListener('click', (e) => {
      e.stopPropagation(); // 阻止事件冒泡到拖拽事件
      this.removeInvoice(invoice.id);
    });
    
    this.invoicesList.appendChild(invoiceEl);
  }

  removeInvoice(id) {
    this.invoices = this.invoices.filter(invoice => invoice.id !== id);
    this.updateInvoiceCount();
    
    // 重新渲染发票列表
    this.invoicesList.innerHTML = '';
    this.invoices.forEach(invoice => this.renderInvoiceItem(invoice));
  }

  updateInvoiceCount() {
    this.invoiceCount.textContent = `(${this.invoices.length})`;
  }

  // 初始化拖拽功能
  initDragAndDrop() {
    let draggedItem = null;
    let draggedIndex = -1;
    
    // 为每个发票项添加拖拽事件
    this.invoicesList.addEventListener('dragstart', (e) => {
      draggedItem = e.target.closest('.invoice-item');
      if(draggedItem) {
        // 记录被拖拽元素的索引
        const items = Array.from(this.invoicesList.querySelectorAll('.invoice-item'));
        draggedIndex = items.indexOf(draggedItem);
        
        // 应用拖拽样式
        setTimeout(() => {
          draggedItem.classList.add('dragging');
        }, 0);
        
        // 设置拖拽效果
        e.dataTransfer.effectAllowed = 'move';
      }
    });
    
    this.invoicesList.addEventListener('dragend', (e) => {
      if(draggedItem) {
        // 清理拖拽样式
        draggedItem.classList.remove('dragging');
        draggedItem = null;
        draggedIndex = -1;
      }
    });
    
    // 允许在列表上拖动
    this.invoicesList.addEventListener('dragover', (e) => {
      e.preventDefault(); // 必须阻止默认行为才能允许放置
      e.dataTransfer.dropEffect = 'move';
    });
    
    this.invoicesList.addEventListener('dragenter', (e) => {
      const targetItem = e.target.closest('.invoice-item');
      if(targetItem && targetItem !== draggedItem) {
        // 添加视觉反馈
        targetItem.style.backgroundColor = '#e3f2fd';
        targetItem.style.transform = 'scale(1.02)';
        targetItem.style.transition = 'all 0.2s ease';
      }
    });
    
    this.invoicesList.addEventListener('dragleave', (e) => {
      const targetItem = e.target.closest('.invoice-item');
      if(targetItem) {
        // 移除视觉反馈
        targetItem.style.backgroundColor = '';
        targetItem.style.transform = 'scale(1)';
      }
    });
    
    this.invoicesList.addEventListener('drop', (e) => {
      e.preventDefault();
      const dropTarget = e.target.closest('.invoice-item');
      
      if(dropTarget && draggedItem && dropTarget !== draggedItem) {
        // 移除视觉反馈
        dropTarget.style.backgroundColor = '';
        dropTarget.style.transform = 'scale(1)';
        
        // 获取所有发票项
        const items = Array.from(this.invoicesList.querySelectorAll('.invoice-item'));
        const dropIndex = items.indexOf(dropTarget);
        
        // 重新排序本地数组
        if(draggedIndex !== -1 && dropIndex !== -1 && draggedIndex !== dropIndex) {
          const movedInvoice = this.invoices[draggedIndex];
          
          // 从原位置移除
          this.invoices.splice(draggedIndex, 1);
          // 插入到新位置
          this.invoices.splice(dropIndex, 0, movedInvoice);
          
          // 重新渲染发票列表以反映新顺序
          this.invoicesList.innerHTML = '';
          this.invoices.forEach(invoice => this.renderInvoiceItem(invoice));
        }
      } else if(dropTarget) {
        dropTarget.style.backgroundColor = '';
        dropTarget.style.transform = 'scale(1)';
      }
    });
  }

  async generatePreview() {
    if (this.invoices.length === 0) {
      alert('请先上传发票文件');
      return;
    }
    
    this.showLoading(true);
    
    try {
      // 获取纸张尺寸
      const paperSize = this.getPaperDimensions();
      
      // 根据每页发票数计算总页数
      const invoicesPerSheet = this.layoutConfig.invoicesPerSheet;
      const totalPages = Math.ceil(this.invoices.length / invoicesPerSheet);
      
      // 生成预览页面数组
      this.previewPages = [];
      for (let i = 0; i < totalPages; i++) {
        // 创建canvas用于生成预览
        const canvas = document.createElement('canvas');
        // 设置较高的分辨率以确保清晰度
        const dpi = 144; // 提高DPI以获得更好的清晰度
        const pxWidth = Math.round(paperSize.width * dpi / 25.4); // mm转px
        const pxHeight = Math.round(paperSize.height * dpi / 25.4);
        
        canvas.width = pxWidth;
        canvas.height = pxHeight;
        const ctx = canvas.getContext('2d');
        
        // 设置白色背景
        ctx.fillStyle = 'white';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        
        // 获取当前页面的发票
        const startIndex = i * invoicesPerSheet;
        const endIndex = Math.min(startIndex + invoicesPerSheet, this.invoices.length);
        const currentPageInvoices = this.invoices.slice(startIndex, endIndex);
        
        // 绘制当前页面的发票
        await this.drawInvoicesOnCanvas(ctx, canvas.width, canvas.height, currentPageInvoices);
        
        // 绘制裁剪线
        this.drawCutLines(ctx, canvas.width, canvas.height);
        
        this.previewPages.push(canvas);
      }
      
      // 显示第一页
      this.currentPageIndex = 0;
      
      // 更新页面计数显示
      this.totalPagesSpan.textContent = this.previewPages.length;
      this.currentPageSpan.textContent = this.currentPageIndex + 1;
      
      // 显示全屏预览
      this.fullscreenPreview.style.display = 'flex';
      
      // 立即更新预览显示
      this.updatePreviewDisplay();
      
    } catch (error) {
      console.error('生成预览失败:', error);
      alert('生成预览失败，请重试');
    } finally {
      this.showLoading(false);
    }
  }

  getPaperDimensions() {
    const orientation = this.paperConfig.orientation;
    
    switch (this.paperConfig.type) {
      case 'A4':
        return orientation === 'portrait' 
          ? { width: 210, height: 297 } 
          : { width: 297, height: 210 };
      case 'A3':
        return orientation === 'portrait' 
          ? { width: 297, height: 420 } 
          : { width: 420, height: 297 };
      case 'letter':
        return orientation === 'portrait' 
          ? { width: 216, height: 279 } 
          : { width: 279, height: 216 };
      default:
        return orientation === 'portrait' 
          ? { width: 210, height: 297 } 
          : { width: 297, height: 210 };
    }
  }

  async drawInvoicesOnCanvas(ctx, canvasWidth, canvasHeight, invoices) {
    const invoicesPerSheet = invoices.length; // 使用传入的发票数组长度
    
    // 计算发票区域
    let rows, cols;
    
    switch (invoicesPerSheet) {
      case 1:
        rows = 1;
        cols = 1;
        break;
      case 2:
        rows = 2;
        cols = 1;
        break;
      case 3:
        rows = 3;
        cols = 1;
        break;
      case 4:
        rows = 2;
        cols = 2;
        break;
      case 6:
        rows = 3;
        cols = 2;
        break;
      default:
        rows = 1;
        cols = 1;
    }
    
    // 移除边距，使用整个页面
    const margin = 0;
    const cellWidth = canvasWidth / cols;
    const cellHeight = canvasHeight / rows;
    
    // 清除画布
    ctx.fillStyle = 'white';
    ctx.fillRect(0, 0, canvasWidth, canvasHeight);
    
    // 绘制每个发票
    for (let i = 0; i < invoices.length; i++) {
      const row = Math.floor(i / cols);
      const col = i % cols;
      
      const x = col * cellWidth;
      const y = row * cellHeight;
      
      const invoice = invoices[i];
      
      // 确保发票有有效的dataUrl
      if (invoice && invoice.dataUrl) {
        try {
          await this.drawImageToCell(ctx, invoice.dataUrl, x, y, cellWidth, cellHeight);
        } catch (error) {
          console.error(`绘制发票 ${invoice.name} 时发生错误:`, error);
          
          // 绘制错误占位符
          ctx.fillStyle = '#f0f0f0';
          ctx.fillRect(x, y, cellWidth, cellHeight);
          
          ctx.strokeStyle = '#ccc';
          ctx.lineWidth = 1;
          ctx.strokeRect(x, y, cellWidth, cellHeight);
          
          ctx.fillStyle = '#999';
          ctx.font = '12px Arial';
          ctx.textAlign = 'center';
          ctx.textBaseline = 'middle';
          ctx.fillText('图像渲染失败', x + cellWidth/2, y + cellHeight/2 - 8);
          ctx.fillText('(请检查文件)', x + cellWidth/2, y + cellHeight/2 + 8);
        }
      } else {
        // 如果发票没有有效的dataUrl，绘制占位符
        ctx.fillStyle = '#f8f8f8';
        ctx.fillRect(x, y, cellWidth, cellHeight);
        
        ctx.strokeStyle = '#ddd';
        ctx.lineWidth = 1;
        ctx.strokeRect(x, y, cellWidth, cellHeight);
        
        ctx.fillStyle = '#aaa';
        ctx.font = '12px Arial';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText('无有效数据', x + cellWidth/2, y + cellHeight/2);
      }
    }
  }

  async drawImageToCell(ctx, imageUrl, x, y, cellWidth, cellHeight) {
    return new Promise((resolve) => {
      const img = new Image();
      img.crossOrigin = 'anonymous';
      
      img.onload = () => {
        // 计算图像在单元格中的位置，保持宽高比并居中
        // 使用稍微小一点的比例（0.95）以确保内容完全可见
        const scaleMargin = 0.95; // 小小的安全边距
        const imgAspect = img.width / img.height;
        const cellAspect = cellWidth / cellHeight;
        
        let renderWidth, renderHeight, renderX, renderY;
        
        if (imgAspect > cellAspect) {
          // 图像更宽，按宽度缩放，垂直居中
          renderWidth = cellWidth * scaleMargin;
          renderHeight = renderWidth / imgAspect;
          renderX = x + (cellWidth - renderWidth) / 2;
          renderY = y + (cellHeight - renderHeight) / 2;
        } else {
          // 图像更高，按高度缩放，水平居中
          renderHeight = cellHeight * scaleMargin;
          renderWidth = renderHeight * imgAspect;
          renderX = x + (cellWidth - renderWidth) / 2;
          renderY = y + (cellHeight - renderHeight) / 2;
        }
        
        // 设置高质量渲染
        ctx.imageSmoothingEnabled = true;
        ctx.imageSmoothingQuality = 'high';
        
        // 在绘制前设置高质量参数
        ctx.mozImageSmoothingEnabled = true;
        ctx.webkitImageSmoothingEnabled = true;
        ctx.msImageSmoothingEnabled = true;
        
        ctx.drawImage(img, renderX, renderY, renderWidth, renderHeight);
        resolve();
      };
      
      img.onerror = () => {
        // 如果图像加载失败，绘制占位符
        ctx.fillStyle = '#f0f0f0';
        ctx.fillRect(x, y, cellWidth, cellHeight);
        
        ctx.strokeStyle = '#ccc';
        ctx.lineWidth = 1;
        ctx.strokeRect(x, y, cellWidth, cellHeight);
        
        ctx.fillStyle = '#999';
        ctx.font = '10px Arial';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText('图像加载失败', x + cellWidth/2, y + cellHeight/2);
        ctx.fillText('(请检查文件)', x + cellWidth/2, y + cellHeight/2 + 15);
        resolve();
      };
      
      img.src = imageUrl;
    });
  }

  drawCutLines(ctx, canvasWidth, canvasHeight) {
    const invoicesPerSheet = this.layoutConfig.invoicesPerSheet;
    
    let rows, cols;
    
    switch (invoicesPerSheet) {
      case 1:
        rows = 1;
        cols = 1;
        break;
      case 2:
        rows = 2;
        cols = 1;
        break;
      case 3:
        rows = 3;
        cols = 1;
        break;
      case 4:
        rows = 2;
        cols = 2;
        break;
      case 6:
        rows = 3;
        cols = 2;
        break;
      default:
        rows = 1;
        cols = 1;
    }
    
    // 设置虚线样式
    ctx.setLineDash([5, 5]);
    ctx.strokeStyle = '#999';
    ctx.lineWidth = 1;
    
    // 计算单元格尺寸（与drawInvoicesOnCanvas中相同的计算方式）
    const cellWidth = canvasWidth / cols;
    const cellHeight = canvasHeight / rows;
    
    // 绘制所有列之间的虚线
    for (let c = 1; c < cols; c++) {
      const x = c * cellWidth;
      ctx.beginPath();
      ctx.moveTo(x, 0);
      ctx.lineTo(x, canvasHeight);
      ctx.stroke();
    }
    
    // 绘制所有行之间的虚线
    for (let r = 1; r < rows; r++) {
      const y = r * cellHeight;
      ctx.beginPath();
      ctx.moveTo(0, y);
      ctx.lineTo(canvasWidth, y);
      ctx.stroke();
    }
    
    // 恢复实线
    ctx.setLineDash([]);
  }

  // 更新预览显示
  updatePreviewDisplay() {
    // 检查DOM元素是否都存在
    if (!this.previewCanvas || !this.currentPageSpan || !this.prevPageBtn || !this.nextPageBtn) {
      console.error('预览所需的DOM元素未找到');
      return;
    }
    
    if (this.previewPages.length === 0) {
      console.warn('没有预览页面可以显示');
      return;
    }
    
    if (this.currentPageIndex >= this.previewPages.length) {
      console.warn('当前页面索引超出范围');
      return;
    }
    
    const currentCanvas = this.previewPages[this.currentPageIndex];
    
    // 将预览页面的内容复制到显示画布
    const displayCanvas = this.previewCanvas;
    const displayCtx = displayCanvas.getContext('2d');
    
    // 确保容器存在
    const container = document.querySelector('.Preview-body');
    if (!container) {
      console.error('找不到预览容器 .Preview-body');
      return;
    }
    
    // 强制重绘容器以确保尺寸正确
    const containerRect = container.getBoundingClientRect();
    const containerWidth = containerRect.width;
    const containerHeight = containerRect.height;
    
    // 最大可用空间，减去一些内边距
    const maxHeight = containerHeight * 0.9;
    const maxWidth = containerWidth * 0.9;
    
    const aspectRatio = currentCanvas.width / currentCanvas.height;
    
    let displayWidth, displayHeight;
    
    // 计算合适的显示尺寸
    if (currentCanvas.width > maxWidth || currentCanvas.height > maxHeight) {
      // 如果原始尺寸超过容器尺寸，则按比例缩小
      if (maxWidth / aspectRatio <= maxHeight) {
        displayWidth = maxWidth;
        displayHeight = maxWidth / aspectRatio;
      } else {
        displayHeight = maxHeight;
        displayWidth = maxHeight * aspectRatio;
      }
    } else {
      // 如果原始尺寸小于容器尺寸，则使用原始尺寸
      displayWidth = currentCanvas.width;
      displayHeight = currentCanvas.height;
    }
    
    // 设置显示画布的尺寸
    displayCanvas.width = displayWidth;
    displayCanvas.height = displayHeight;
    
    // 绘制缩放后的图像
    displayCtx.drawImage(
      currentCanvas,
      0, 0, currentCanvas.width, currentCanvas.height,
      0, 0, displayWidth, displayHeight
    );
    
    // 更新页面计数显示
    this.currentPageSpan.textContent = this.currentPageIndex + 1;
    
    // 更新翻页按钮状态
    this.prevPageBtn.disabled = this.currentPageIndex === 0;
    this.nextPageBtn.disabled = this.currentPageIndex === this.previewPages.length - 1;
  }

  // 下一页
  nextPage() {
    if (this.currentPageIndex < this.previewPages.length - 1) {
      this.currentPageIndex++;
      this.updatePreviewDisplay();
    }
  }

  // 上一页
  previousPage() {
    if (this.currentPageIndex > 0) {
      this.currentPageIndex--;
      this.updatePreviewDisplay();
    }
  }

  async exportPDF() {
    if (this.invoices.length === 0) {
      alert('请先上传发票文件');
      return;
    }
    
    // 临时禁用按钮并显示状态
    this.exportBtn.disabled = true;
    this.exportBtn.textContent = '导出中...';
    
    this.showLoading(true);
    
    try {
      // 使用jsPDF创建PDF
      const { jsPDF } = window.jspdf;
      
      // 获取纸张尺寸
      const paperSize = this.getPaperDimensions();
      const pdf = new jsPDF({
        orientation: this.paperConfig.orientation,
        unit: 'mm',
        format: [paperSize.width, paperSize.height]
      });
      
      // 每页最多处理指定数量的发票
      const invoicesPerSheet = this.layoutConfig.invoicesPerSheet;
      
      for (let i = 0; i < this.invoices.length; i += invoicesPerSheet) {
        if (i > 0) {
          // 添加新页面
          pdf.addPage([paperSize.width, paperSize.height], this.paperConfig.orientation);
        }
        
        // 当前页面的发票
        const currentBatch = this.invoices.slice(i, i + invoicesPerSheet);
        
        // 计算布局
        let rows, cols;
        
        switch (invoicesPerSheet) {
          case 1:
            rows = 1;
            cols = 1;
            break;
          case 2:
            rows = 2;
            cols = 1;
            break;
          case 3:
            rows = 3;
            cols = 1;
            break;
          case 4:
            rows = 2;
            cols = 2;
            break;
          case 6:
            rows = 3;
            cols = 2;
            break;
          default:
            rows = 1;
            cols = 1;
        }
        
        // 移除边距，使用整个页面
        const margin = 0; // 移除边距
        const cellWidth = paperSize.width / cols;
        const cellHeight = paperSize.height / rows;
        
        // 在当前页面绘制发票
        for (let j = 0; j < currentBatch.length; j++) {
          const row = Math.floor(j / cols);
          const col = j % cols;
          
          const x = col * cellWidth;
          const y = row * cellHeight;
          
          const invoice = currentBatch[j];
          
          // 等待图像加载完成
          await new Promise((resolve) => {
            const img = new Image();
            img.setAttribute('crossOrigin', 'anonymous');
            
            img.onload = async () => {
              // 计算图像在PDF中的尺寸，保持宽高比并居中
              // 使用稍微小一点的比例（0.95）以确保内容完全可见
              const scaleMargin = 0.95; // 小小的安全边距
              
              const imgAspect = img.width / img.height;
              const cellAspect = cellWidth / cellHeight;
              
              let renderWidth, renderHeight, renderX, renderY;
              
              if (imgAspect > cellAspect) {
                // 图像更宽，按宽度缩放，垂直居中
                renderWidth = cellWidth * scaleMargin;
                renderHeight = renderWidth / imgAspect;
                renderX = x + (cellWidth - renderWidth) / 2;
                renderY = y + (cellHeight - renderHeight) / 2;
              } else {
                // 图像更高，按高度缩放，水平居中
                renderHeight = cellHeight * scaleMargin;
                renderWidth = renderHeight * imgAspect;
                renderX = x + (cellWidth - renderWidth) / 2;
                renderY = y + (cellHeight - renderHeight) / 2;
              }
              
              // 将图像添加到PDF
              try {
                // 使用PNG格式，保留所有细节，质量设置为最高
                pdf.addImage(
                  invoice.dataUrl,
                  'PNG',  // 使用PNG格式以保留透明度和细节
                  renderX,
                  renderY,
                  renderWidth,
                  renderHeight,
                  undefined,
                  'SLOW'  // 最高质量，最慢的压缩算法
                );
              } catch (e) {
                console.error('添加PNG图像到PDF失败:', e);
                // 如果PNG失败，尝试使用JPEG格式
                try {
                  pdf.addImage(
                    invoice.dataUrl,
                    'JPEG',
                    renderX,
                    renderY,
                    renderWidth,
                    renderHeight,
                    undefined,
                    'SLOW'  // 最高质量
                  );
                } catch (e2) {
                  console.error('使用JPEG格式添加图像也失败:', e2);
                }
              }
              
              resolve();
            };
            
            img.onerror = (err) => {
              console.error('加载图像失败:', err);
              resolve(); // 继续处理下一个发票
            };
            
            img.src = invoice.dataUrl;
          });
        }
        
        // 只绘制实际有发票的区域的分割线
        const actualInvoiceCount = currentBatch.length;
        
        // 绘制裁剪线
        pdf.setDrawColor(150); // 灰色
        pdf.setLineWidth(0.1);
        pdf.setLineDash([0.5, 0.5]); // 设置虚线
        
        // 计算单元格尺寸（与drawInvoicesOnCanvas中相同的计算方式）
        const calculatedCellWidth = paperSize.width / cols;
        const calculatedCellHeight = paperSize.height / rows;
        
        // 绘制所有列之间的虚线
        for (let c = 1; c < cols; c++) {
          const x = c * calculatedCellWidth;
          pdf.line(x, 0, x, paperSize.height);
        }
        
        // 绘制所有行之间的虚线
        for (let r = 1; r < rows; r++) {
          const y = r * calculatedCellHeight;
          pdf.line(0, y, paperSize.width, y);
        }
        
        // 恢复实线
        pdf.setLineDash([0, 0]);
      }
      
      // 生成文件名
      const fileName = this.generateFileName();
      
      // 保存PDF
      pdf.save(fileName);
      
      alert(`PDF导出成功！文件已保存为：${fileName}\n\n共${this.invoices.length}张发票已合并导出。`);
    } catch (error) {
      console.error('导出PDF失败:', error);
      alert('导出PDF失败，请重试\n错误详情: ' + error.message);
    } finally {
      this.showLoading(false);
      // 恢复按钮状态
      this.exportBtn.disabled = false;
      this.exportBtn.textContent = '导出发票PDF';
    }
  }

  generateFileName() {
    const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const count = this.invoices.length;
    
    // 获取纸张尺寸描述
    let paperSizeDesc = '';
    switch (this.paperConfig.type) {
      case 'A4':
        paperSizeDesc = this.paperConfig.orientation === 'portrait' ? 'A4_Portrait' : 'A4_Landscape';
        break;
      case 'A3':
        paperSizeDesc = this.paperConfig.orientation === 'portrait' ? 'A3_Portrait' : 'A3_Landscape';
        break;
      case 'letter':
        paperSizeDesc = this.paperConfig.orientation === 'portrait' ? 'Letter_Portrait' : 'Letter_Landscape';
        break;
      default:
        paperSizeDesc = 'A4_Portrait';
    }
    
    // 计算总金额（如果有）
    let totalAmount = 0;
    // 这里可以添加从发票中提取金额的逻辑
    
    // 格式化项目周期日期
    let formattedProjectCycleStart = '';
    let formattedProjectCycleEnd = '';
    if (this.namingConfig.projectCycleStart) {
      const startDate = new Date(this.namingConfig.projectCycleStart);
      formattedProjectCycleStart = startDate.toISOString().slice(0, 10).replace(/-/g, '');
    }
    
    if (this.namingConfig.projectCycleEnd) {
      const endDate = new Date(this.namingConfig.projectCycleEnd);
      formattedProjectCycleEnd = endDate.toISOString().slice(0, 10).replace(/-/g, '');
    }
    
    // 使用固定格式：项目名称+产品+项目周期
    let fileName = '';
    if (this.namingConfig.projectName) {
      fileName += this.namingConfig.projectName;
    }
    
    if (this.namingConfig.product) {
      fileName += fileName ? '_' + this.namingConfig.product : this.namingConfig.product;
    }
    
    if (formattedProjectCycleStart && formattedProjectCycleEnd) {
      fileName += fileName ? '_' + formattedProjectCycleStart + '-' + formattedProjectCycleEnd : 
                            formattedProjectCycleStart + '-' + formattedProjectCycleEnd;
    } else if (formattedProjectCycleStart) {
      fileName += fileName ? '_' + formattedProjectCycleStart : formattedProjectCycleStart;
    }
    
    // 如果没有填写任何自定义名称，则使用默认格式
    if (!fileName) {
      fileName = `发票_${dateStr}_${count}张`;
    }
    
    return fileName + '.pdf';
  }

  showLoading(show) {
    this.loadingOverlay.style.display = show ? 'flex' : 'none';
  }
}

// 应用初始化
document.addEventListener('DOMContentLoaded', () => {
  new InvoiceManager();
});