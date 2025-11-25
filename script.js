'use strict';

const APPS_SCRIPT_URL = (window.APPS_SCRIPT_URL || '').trim();
const SEDES = ['SL', 'LPG', 'SC', 'SCH', 'PB-2', 'E PB-2', 'LG', 'VM', 'BC'];
const STORAGE_KEY = 'latata-catalog-v1';

const state = {
  products: [],
};

const elements = {
  navButtons: () => document.querySelectorAll('.nav-btn'),
  viewTriggers: () => document.querySelectorAll('[data-view-target]'),
  views: () => document.querySelectorAll('.view'),
  solicitudRows: document.getElementById('solicitud-product-rows'),
  addSolicitudRowBtn: document.getElementById('add-product-row'),
  registroRows: document.getElementById('registros-product-rows'),
  addRegistroRowBtn: document.getElementById('add-registro-row'),
  mermaRows: document.getElementById('merma-product-rows'),
  addMermaRowBtn: document.getElementById('add-merma-row'),
  catalogBody: document.getElementById('catalog-body'),
  catalogStatus: document.getElementById('catalog-status'),
  catalogSearch: document.getElementById('catalog-search'),
  refreshCatalogBtn: document.getElementById('refresh-catalog'),
  toast: document.getElementById('toast'),
  envWarning: document.getElementById('env-warning'),
  productOptions: document.getElementById('product-options'),
};

const queryAll = (scope, selector) => {
  if (!scope) return [];
  if (typeof scope.querySelectorAll === 'function') {
    return scope.querySelectorAll(selector);
  }
  return [];
};

init();

function init() {
  setupNavigation();
  populateSedeSelects();
  initSolicitudesForm();
  initRegistrosForm();
  initMermaForm();
  setupProductCombos();
  syncProductCombosState();
  setupSingleProductHintButtons();
  initCatalogView();
  loadCatalogFromCache();
  fetchProducts();
  toggleEnvWarning(!APPS_SCRIPT_URL);
}

function setupNavigation() {
  elements.viewTriggers().forEach((trigger) => {
    trigger.addEventListener('click', () => showView(trigger.dataset.viewTarget));
  });
}

function showView(target) {
  if (!target) return;
  elements.views().forEach((view) => {
    view.classList.toggle('active', view.dataset.view === target);
  });
  elements.navButtons().forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.viewTarget === target);
  });
}

function populateSedeSelects() {
  document.querySelectorAll('[data-role="sede-select"]').forEach((select) => {
    const currentValue = select.value;
    const options = [
      '<option value="" disabled selected>Selecciona una sede</option>',
      ...SEDES.map((sede) => `<option value="${sede}">${sede}</option>`),
    ].join('');
    select.innerHTML = options;
    if (currentValue && SEDES.includes(currentValue)) {
      select.value = currentValue;
    }
  });
}

function initSolicitudesForm() {
  const form = document.getElementById('solicitudes-form');
  if (!form || !elements.solicitudRows) return;

  resetSolicitudRows();
  elements.addSolicitudRowBtn?.addEventListener('click', () => addSolicitudRow());

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    if (typeof form.reportValidity === 'function' && !form.reportValidity()) {
      return;
    }
    const formData = new FormData(form);
    let items;

    try {
      items = collectSolicitudItems();
    } catch (error) {
      showToast(error.message || 'Corrige las cantidades ingresadas.', 'error');
      return;
    }

    if (!items.length) {
      showToast('Agrega al menos un producto.', 'error');
      return;
    }

    if (hasDuplicateProducts(items)) {
      showToast('No puedes repetir un producto en la solicitud.', 'error');
      return;
    }

    const payload = {
      fecha: formData.get('fecha') || '',
      hora: formData.get('hora') || '',
      sede: formData.get('sede') || '',
      responsable: formData.get('responsable') || '',
      items,
    };

    try {
      toggleFormLoading(form, true);
      await postData('createSolicitud', payload);
      showToast('Solicitud registrada correctamente.', 'success');
      form.reset();
      resetSolicitudRows();
    } catch (error) {
      showToast(error.message || 'Error al registrar la solicitud.', 'error');
    } finally {
      toggleFormLoading(form, false);
    }
  });
}

function collectSolicitudItems() {
  return collectItems(elements.solicitudRows, 'cantidad solicitada', (product, quantity) => ({
    code: product.code,
    description: product.description,
    unit: product.unit,
    quantity,
  }));
}

function resetSolicitudRows() {
  if (!elements.solicitudRows) return;
  elements.solicitudRows.innerHTML = '';
  addSolicitudRow();
}

function addSolicitudRow() {
  if (!elements.solicitudRows) return;
  const row = document.createElement('div');
  row.className = 'product-row';
  row.innerHTML = `
    <label>
      <span>Producto</span>
      <input type="hidden" data-role="product-value" />
      <input
        type="text"
        class="product-combo"
        data-role="product-combo"
        placeholder="Seleccione o escriba un producto"
        list="product-options"
        autocomplete="off"
        required
      />
      <small class="unit-hint" data-unit-output>Unidad: --</small>
    </label>
    <label>
      <span>Cantidad solicitada</span>
      <input type="number" min="0" step="1" value="1" required />
    </label>
    <div class="product-row__actions">
      <span class="unit-badge" data-unit>--</span>
      <button type="button" class="remove-row">Eliminar</button>
    </div>
  `;

  elements.solicitudRows.appendChild(row);
  setupProductCombos(row);
  row.querySelector('.remove-row').addEventListener('click', () =>
    removeProductRow(row, elements.solicitudRows)
  );
  updateRowUnit(row);
}

function removeProductRow(row, container) {
  const target = container || row.parentElement;
  if (!target) return;
  if (target.children.length === 1) {
    showToast('Debes mantener al menos un producto.', 'error');
    return;
  }
  row.remove();
}

function initRegistrosForm() {
  const form = document.getElementById('registros-form');
  if (!form || !elements.registroRows) return;

  resetRegistroRows();
  elements.addRegistroRowBtn?.addEventListener('click', () => addRegistroRow());

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    if (typeof form.reportValidity === 'function' && !form.reportValidity()) {
      return;
    }
    const formData = new FormData(form);
    let items;

    try {
      items = collectRegistroItems();
    } catch (error) {
      showToast(error.message || 'Corrige las cantidades ingresadas.', 'error');
      return;
    }

    if (!items.length) {
      showToast('Agrega al menos un producto.', 'error');
      return;
    }

    if (hasDuplicateProducts(items)) {
      showToast('Cada producto solo puede aparecer una vez por registro.', 'error');
      return;
    }

    const payload = {
      fecha: formData.get('fecha') || '',
      hora: formData.get('hora') || '',
      sede: formData.get('sede') || '',
      responsableEntrega: formData.get('responsableEntrega') || '',
      sinSolicitud: formData.get('sinSolicitud') === 'on',
      items,
    };

    try {
      toggleFormLoading(form, true);
      await postData('recordEntrega', payload);
      showToast('Registros procesados.', 'success');
      form.reset();
      resetRegistroRows();
    } catch (error) {
      showToast(error.message || 'Error al registrar la entrega.', 'error');
    } finally {
      toggleFormLoading(form, false);
    }
  });
}

function addRegistroRow() {
  if (!elements.registroRows) return;
  const row = document.createElement('div');
  row.className = 'product-row';
  row.innerHTML = `
    <label>
      <span>Producto</span>
      <input type="hidden" data-role="product-value" />
      <input
        type="text"
        class="product-combo"
        data-role="product-combo"
        placeholder="Seleccione o escriba un producto"
        list="product-options"
        autocomplete="off"
        required
      />
      <small class="unit-hint" data-unit-output>Unidad: --</small>
    </label>
    <label>
      <span>Cantidad entregada</span>
      <input type="number" min="0" step="1" value="1" required />
    </label>
    <div class="product-row__actions">
      <span class="unit-badge" data-unit>--</span>
      <button type="button" class="remove-row">Eliminar</button>
    </div>
  `;

  elements.registroRows.appendChild(row);
  setupProductCombos(row);
  row.querySelector('.remove-row').addEventListener('click', () =>
    removeProductRow(row, elements.registroRows)
  );
  updateRowUnit(row);
}

function resetRegistroRows() {
  if (!elements.registroRows) return;
  elements.registroRows.innerHTML = '';
  addRegistroRow();
}

function collectRegistroItems() {
  return collectItems(elements.registroRows, 'cantidad entregada', (product, quantity) => ({
    productCode: product.code,
    productName: product.description,
    unit: product.unit,
    cantidadEntregada: quantity,
  }));
}

function initMermaForm() {
  const form = document.getElementById('merma-form');
  if (!form || !elements.mermaRows) return;

  resetMermaRows();
  elements.addMermaRowBtn?.addEventListener('click', () => addMermaRow());

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    if (typeof form.reportValidity === 'function' && !form.reportValidity()) {
      return;
    }
    const formData = new FormData(form);
    let items;

    try {
      items = collectMermaItems();
    } catch (error) {
      showToast(error.message || 'Corrige las cantidades ingresadas.', 'error');
      return;
    }

    if (!items.length) {
      showToast('Agrega al menos un producto.', 'error');
      return;
    }

    if (hasDuplicateProducts(items)) {
      showToast('No repitas productos en la merma.', 'error');
      return;
    }

    const payload = {
      fecha: formData.get('fecha') || '',
      hora: formData.get('hora') || '',
      sede: formData.get('sede') || 'BC',
      responsable: formData.get('responsable') || '',
      items,
    };

    try {
      toggleFormLoading(form, true);
      await postData('recordMerma', payload);
      showToast('Mermas registradas para BC.', 'success');
      form.reset();
      resetMermaRows();
    } catch (error) {
      showToast(error.message || 'Error al registrar la merma.', 'error');
    } finally {
      toggleFormLoading(form, false);
    }
  });
}

function addMermaRow() {
  if (!elements.mermaRows) return;
  const row = document.createElement('div');
  row.className = 'product-row';
  row.innerHTML = `
    <label>
      <span>Producto</span>
      <input type="hidden" data-role="product-value" />
      <input
        type="text"
        class="product-combo"
        data-role="product-combo"
        placeholder="Seleccione o escriba un producto"
        list="product-options"
        autocomplete="off"
        required
      />
      <small class="unit-hint" data-unit-output>Unidad: --</small>
    </label>
    <label>
      <span>Cantidad merma</span>
      <input type="number" min="0" step="1" value="1" required />
    </label>
    <div class="product-row__actions">
      <span class="unit-badge" data-unit>--</span>
      <button type="button" class="remove-row">Eliminar</button>
    </div>
  `;

  elements.mermaRows.appendChild(row);
  setupProductCombos(row);
  row.querySelector('.remove-row').addEventListener('click', () =>
    removeProductRow(row, elements.mermaRows)
  );
  updateRowUnit(row);
}

function resetMermaRows() {
  if (!elements.mermaRows) return;
  elements.mermaRows.innerHTML = '';
  addMermaRow();
}

function collectMermaItems() {
  return collectItems(elements.mermaRows, 'cantidad de merma', (product, quantity) => ({
    productCode: product.code,
    productName: product.description,
    unit: product.unit,
    cantidadMerma: quantity,
  }));
}

function setupProductCombos(scope = document) {
  const inputs = queryAll(scope, '[data-role="product-combo"]');
  inputs.forEach((input) => {
    if (!input || input.dataset.comboBound === 'true') return;
    const hiddenInput = input.parentElement?.querySelector('[data-role="product-value"]');
    const unitOutput = input.parentElement?.querySelector('[data-unit-output]');
    const getBadge = () => input.closest('.product-row')?.querySelector('[data-unit]');

    const clearSelection = () => {
      delete input.dataset.code;
      if (hiddenInput) hiddenInput.value = '';
      if (unitOutput) unitOutput.textContent = 'Unidad: --';
      const badge = getBadge();
      if (badge) badge.textContent = '--';
    };

    const commitSelection = () => {
      const product = getProductFromInput(input);
      if (!product) {
        clearSelection();
        return null;
      }
      input.dataset.code = product.code;
      input.value = formatProductOption(product);
      if (hiddenInput) hiddenInput.value = product.code;
      if (unitOutput) unitOutput.textContent = `Unidad: ${product.unit}`;
      const badge = getBadge();
      if (badge) badge.textContent = product.unit;
      return product;
    };

    input.addEventListener('input', () => {
      clearSelection();
    });

    input.addEventListener('change', commitSelection);

    input.addEventListener('blur', () => {
      if (!commitSelection() && input.value.trim()) {
        showToast('Selecciona un producto del catálogo.', 'error');
        input.value = '';
      }
    });

    input.dataset.comboBound = 'true';
  });
}

function updateRowUnit(row) {
  const input = row.querySelector('[data-role="product-combo"]');
  const product = getProductFromInput(input);
  const badge = row.querySelector('[data-unit]');
  const unitHint = row.querySelector('[data-unit-output]');
  const unitLabel = product?.unit || '--';
  if (badge) badge.textContent = unitLabel;
  if (unitHint) unitHint.textContent = `Unidad: ${unitLabel}`;
}

function getProductFromInput(input) {
  if (!input) return null;
  if (input.dataset.code) {
    const productByDataset = findProduct(input.dataset.code);
    if (productByDataset) return productByDataset;
  }

  const rawValue = input.value?.trim();
  if (!rawValue) return null;
  const [codeCandidate] = rawValue.split(' · ');
  if (codeCandidate) {
    const productByCode = findProduct(codeCandidate.trim());
    if (productByCode) return productByCode;
  }

  return (
    state.products.find((product) => {
      const formatted = formatProductOption(product).toLowerCase();
      return (
        formatted === rawValue.toLowerCase() ||
        product.description.toLowerCase() === rawValue.toLowerCase()
      );
    }) || null
  );
}

function formatProductOption(product) {
  return `${product.code} · ${product.description}`;
}

function updateProductOptionsList() {
  if (!elements.productOptions) return;
  elements.productOptions.innerHTML = state.products
    .map((product) => `<option value="${formatProductOption(product)}"></option>`)
    .join('');
}

function syncProductCombosState() {
  const hasProducts = state.products.length > 0;
  document.querySelectorAll('[data-role="product-combo"]').forEach((input) => {
    input.placeholder = hasProducts
      ? 'Seleccione o escriba un producto'
      : 'Sin catálogo disponible';
    input.disabled = !hasProducts;
    if (hasProducts && input.dataset.code) {
      const product = findProduct(input.dataset.code);
      if (product) {
        input.value = formatProductOption(product);
        const unitOutput = input.parentElement?.querySelector('[data-unit-output]');
        if (unitOutput) unitOutput.textContent = `Unidad: ${product.unit}`;
        const badge = input.closest('.product-row')?.querySelector('[data-unit]');
        if (badge) badge.textContent = product.unit;
      }
    }
  });
}

function refreshProductCombos() {
  updateProductOptionsList();
  syncProductCombosState();
}

function resetProductCombos(scope = document) {
  queryAll(scope, '[data-role="product-combo"]').forEach((input) => {
    delete input.dataset.code;
    input.value = '';
    const hiddenInput = input.parentElement?.querySelector('[data-role="product-value"]');
    if (hiddenInput) hiddenInput.value = '';
    const unitOutput = input.parentElement?.querySelector('[data-unit-output]');
    if (unitOutput) unitOutput.textContent = 'Unidad: --';
    const badge = input.closest('.product-row')?.querySelector('[data-unit]');
    if (badge) badge.textContent = '--';
  });
}

function hasDuplicateProducts(items) {
  const seen = new Set();
  for (const item of items) {
    const code = item.code || item.productCode;
    if (!code) {
      continue;
    }
    if (seen.has(code)) {
      return true;
    }
    seen.add(code);
  }
  return false;
}

function setupSingleProductHintButtons() {
  document.querySelectorAll('[data-prefill-product]').forEach((button) => {
    button.addEventListener('click', () => {
      const target = document.querySelector(button.dataset.prefillProduct);
      if (target) {
        target.focus();
        showToast('Selecciona un producto del catálogo.', 'info');
      }
    });
  });
}

function collectItems(container, quantityLabel, mapper) {
  const rows = container?.querySelectorAll('.product-row');
  if (!rows) return [];

  const items = [];
  for (const row of rows) {
    const combo = row.querySelector('[data-role="product-combo"]');
    const quantityInput = row.querySelector('input[type="number"]');
    const product = getProductFromInput(combo);
    if (!product) {
      combo?.focus?.();
      throw new Error('Selecciona un producto del catálogo en cada fila.');
    }

    const quantity = parseIntegerQuantity(quantityInput?.value);
    if (quantity === null) {
      quantityInput?.focus();
      throw new Error(`La ${quantityLabel} debe ser un número entero mayor o igual a 0.`);
    }

    items.push(mapper(product, quantity));
  }

  return items;
}

function initCatalogView() {
  elements.catalogSearch?.addEventListener('input', (event) => {
    const term = event.target.value.trim().toLowerCase();
    const filtered = state.products.filter(
      (product) =>
        product.code.toLowerCase().includes(term) ||
        product.description.toLowerCase().includes(term)
    );
    renderCatalog(filtered);
  });

  elements.refreshCatalogBtn?.addEventListener('click', () => fetchProducts(true));
}

function loadCatalogFromCache() {
  try {
    const cached = localStorage.getItem(STORAGE_KEY);
    if (!cached) return;
    const parsed = JSON.parse(cached);
    if (Array.isArray(parsed)) {
      state.products = parsed;
      refreshProductCombos();
      renderCatalog(parsed);
      setCatalogStatus('Catálogo desde caché', false);
    }
  } catch (error) {
    console.warn('No se pudo leer el catálogo en caché.', error);
  }
}

function parseIntegerQuantity(value) {
  const raw = typeof value === 'number' ? value.toString() : String(value ?? '').trim();
  if (raw === '') {
    return null;
  }
  const parsed = Number(raw);
  if (!Number.isFinite(parsed) || !Number.isInteger(parsed) || parsed < 0) {
    return null;
  }
  return parsed;
}

async function fetchProducts(showToastOnSuccess = false) {
  if (!APPS_SCRIPT_URL) {
    setCatalogStatus('Configura la URL del Apps Script.', true);
    return;
  }

  setCatalogStatus('Sincronizando catálogo...', false);

  try {
    const response = await fetch(`${APPS_SCRIPT_URL}?action=getProducts`, { cache: 'no-store' });
    const data = await response.json();
    if (!data.success) {
      throw new Error(data.message || 'No se pudo sincronizar el catálogo.');
    }

    state.products = data?.data?.products || [];
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.products));
    refreshProductCombos();
    renderCatalog(state.products);
    setCatalogStatus('Catálogo actualizado', false);
    if (showToastOnSuccess) {
      showToast('Catálogo sincronizado.', 'success');
    }
  } catch (error) {
    console.error(error);
    setCatalogStatus(error.message, true);
    showToast(error.message, 'error');
  }
}

function renderCatalog(products) {
  if (!elements.catalogBody) return;
  if (!products.length) {
    elements.catalogBody.innerHTML = '<tr><td colspan="3" class="muted">Sin resultados.</td></tr>';
    return;
  }

  const rows = products
    .map(
      (product) => `
        <tr>
          <td>${product.code}</td>
          <td>${product.description}</td>
          <td>${product.unit}</td>
        </tr>`
    )
    .join('');
  elements.catalogBody.innerHTML = rows;
}

function findProduct(code) {
  return state.products.find((product) => product.code === code);
}

async function postData(action, payload) {
  if (!APPS_SCRIPT_URL) {
    throw new Error('Configura la URL del Apps Script.');
  }

  const response = await fetch(APPS_SCRIPT_URL, {
    method: 'POST',
    body: JSON.stringify({ action, payload }),
  });

  const data = await response.json();
  if (!data.success) {
    throw new Error(data.message || 'Error en la operación.');
  }
  return data;
}

function toggleFormLoading(form, loading) {
  const submitBtn = form.querySelector('button[type="submit"]');
  if (submitBtn) {
    submitBtn.disabled = loading;
  }
}

function setCatalogStatus(message, isError) {
  if (!elements.catalogStatus) return;
  elements.catalogStatus.textContent = message;
  elements.catalogStatus.classList.toggle('error', Boolean(isError));
}

function toggleEnvWarning(show) {
  if (!elements.envWarning) return;
  elements.envWarning.classList.toggle('hidden', !show);
}

let toastTimeout;
function showToast(message, type = 'info') {
  if (!elements.toast) return;
  elements.toast.textContent = message;
  elements.toast.className = `toast show ${type}`;
  clearTimeout(toastTimeout);
  toastTimeout = setTimeout(() => {
    elements.toast.classList.remove('show');
  }, 3500);
}
