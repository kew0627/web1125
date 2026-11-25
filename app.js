'use strict';

const DAILY_ALLOWANCE = 25_000;
const DAILY_MEAL = 25_000;
const LODGING_LIMITS = { seoul: 100_000, metro: 80_000, other: 70_000 };
const DAY_MS = 86_400_000;
const format = new Intl.NumberFormat('ko-KR');
const get = (id) => document.getElementById(id);
const form = get('trip-form');
const lodgingRows = get('lodging-rows');
const error = get('error');
let currentResult = null;

function setTripType() {
  const inside = form.elements['trip-type'].value === 'inside';
  for (const section of [get('inside'), get('outside')]) {
    section.hidden = (section.id === 'inside') !== inside;
    for (const control of section.querySelectorAll('input, select, button')) {
      control.disabled = section.hidden;
    }
  }
  updateCategory();
}

function updateCategory() {
  const first = get('category').value === 'first';
  for (const field of document.querySelectorAll('#region-field, .lodging-region, #excess-field, #excess-hint')) {
    field.hidden = first;
  }
}

function addLodging() {
  const row = get('lodging-template').content.firstElementChild.cloneNode(true);
  row.querySelector('.remove').addEventListener('click', () => {
    row.remove();
    invalidateResult();
    get('add-lodging').focus();
  });
  lodgingRows.append(row);
  updateCategory();
  return row;
}

function tripDays() {
  return (get('return-date').valueAsNumber - get('departure').valueAsNumber) / DAY_MS + 1;
}

function updateDuration() {
  const days = tripDays();
  const hint = get('duration');
  hint.textContent = '출장일과 복귀일을 모두 포함해 계산합니다.';
  if (days > 0) hint.textContent = `출장 ${format.format(days)}일 · 날짜 기준 최대 ${format.format(days - 1)}박`;
  if (days <= 0) hint.textContent = '복귀일은 출장일과 같거나 이후여야 합니다.';
}

function wholeNumber(input) {
  const value = input.value === '' ? 0 : input.valueAsNumber;
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new Error('금액과 일수에는 0 이상의 정수를 입력하세요.');
  }
  return value;
}

function costOfLodging(days, first) {
  let total = 0;
  let nights = 0;
  let reduced = false;
  const region = get('region').value;
  for (const row of lodgingRows.children) {
    const rate = wholeNumber(row.querySelector('.rate'));
    const count = wholeNumber(row.querySelector('.nights'));
    if (rate && !count) throw new Error('1박 실비를 입력한 숙박 항목에 박수를 입력하세요.');
    const rowRegion = row.querySelector('select').value;
    const limit = LODGING_LIMITS[rowRegion === 'default' ? region : rowRegion];
    const eligible = first ? rate : Math.min(rate, get('excess').checked ? limit * 13 / 10 : limit);
    total += eligible * count;
    nights += count;
    reduced ||= count > 0 && eligible < rate;
  }
  if (nights >= days) throw new Error('숙박 박수 합계는 출장일수보다 작아야 합니다.');
  return { total, reduced };
}

function calculateOutside(days) {
  const half = wholeNumber(get('half-days'));
  const noMeal = wholeNumber(get('no-meal-days'));
  const bonusDays = wholeNumber(get('bonus-days'));
  const savings = wholeNumber(get('savings'));
  if (Math.max(half, noMeal, bonusDays) > days) throw new Error('각 적용 일수는 전체 출장일수를 넘을 수 없습니다.');
  if (Boolean(savings) !== Boolean(bonusDays)) throw new Error('공적 항공마일리지를 적용하려면 절약한 운임과 적용 일수를 모두 입력하세요.');

  const lodging = costOfLodging(days, get('category').value === 'first');
  const bonus = Math.min(savings, bonusDays * DAILY_ALLOWANCE / 2);
  const items = [
    ['운임', wholeNumber(get('fare'))],
    ['일비', days * DAILY_ALLOWANCE - half * DAILY_ALLOWANCE / 2],
    ['숙박비', lodging.total],
    ['식비', (days - noMeal) * DAILY_MEAL],
  ];
  if (bonus) items.push(['공적 항공마일리지 추가 일비', bonus]);
  const notes = [];
  if (half) notes.push(`일비 절반 지급 ${format.format(half)}일을 반영했습니다.`);
  if (noMeal) notes.push(`식비 지급 제외 ${format.format(noMeal)}일을 반영했습니다.`);
  if (lodging.reduced) notes.push('1박 지급 상한을 초과한 숙박비를 조정했습니다.');
  return { items, notes, status: `근무지 외 · ${format.format(days)}일` };
}

function calculateInside(days) {
  if (days !== 1) throw new Error('근무지 내 출장은 출장일과 복귀일을 같은 날짜로 입력하세요.');
  const vehicle = get('vehicle').value;
  const base = get('hours').value === 'long' ? 20_000 : 10_000;
  const amount = vehicle === 'assigned' ? 0 : base - (vehicle === 'deduct' ? 10_000 : 0);
  const notes = [];
  if (vehicle === 'deduct') notes.push('공무용 차량 기준에 따라 10,000원을 감액했습니다.');
  if (vehicle === 'assigned') notes.push('전용차량 배정자는 근무지 내 출장 여비 지급 대상에서 제외됩니다.');
  return { items: [['근무지 내 출장 여비', amount]], notes, status: '근무지 내 · 1일' };
}

function renderResult(items, notes, status, total) {
  get('total').firstChild.textContent = total === null ? '— ' : `${format.format(total)} `;
  get('result-status').textContent = status;
  const dates = get('result-dates');
  dates.hidden = total === null;
  dates.textContent = dates.hidden ? '' : `${get('departure').value} ~ ${get('return-date').value}`;
  get('print-result').disabled = total === null;
  get('save-xlsx').disabled = total === null;
  get('export-error').hidden = true;
  get('breakdown').replaceChildren(...items.map(([name, amount]) => {
    const row = document.createElement('div');
    const term = document.createElement('dt');
    const value = document.createElement('dd');
    term.textContent = name;
    value.textContent = amount === null ? '—' : `${format.format(amount)}원`;
    row.append(term, value);
    return row;
  }));
  const list = get('result-notes');
  list.replaceChildren(...notes.map((note) => {
    const item = document.createElement('li');
    item.textContent = note;
    return item;
  }));
  list.hidden = notes.length === 0;
}

function invalidateResult(message = '조건을 입력하고 여비를 계산하세요.') {
  const names = form.elements['trip-type'].value === 'inside' ? ['근무지 내 출장 여비'] : ['운임', '일비', '숙박비', '식비'];
  renderResult(names.map((name) => [name, null]), [], message, null);
  currentResult = null;
  error.hidden = true;
}

form.addEventListener('input', () => {
  setTripType();
  updateDuration();
  const message = currentResult ? '조건이 바뀌었습니다. 다시 계산하세요.' : get('result-status').textContent;
  invalidateResult(message);
});

form.addEventListener('invalid', (event) => {
  const details = event.target.closest('details');
  if (details) details.open = true;
}, true);

form.addEventListener('submit', (event) => {
  event.preventDefault();
  try {
    const days = tripDays();
    if (!Number.isInteger(days) || days < 1) throw new Error('출장일과 복귀일을 순서에 맞게 입력하세요.');
    const result = form.elements['trip-type'].value === 'inside' ? calculateInside(days) : calculateOutside(days);
    const total = result.items.reduce((sum, [, amount]) => sum + amount, 0);
    if (!Number.isSafeInteger(total)) throw new Error('입력 금액이 너무 큽니다.');
    renderResult(result.items, result.notes, result.status, total);
    error.hidden = true;
    currentResult = {
      ...result,
      days,
      total,
      departure: get('departure').value,
      returnDate: get('return-date').value,
      tripType: form.elements['trip-type'].value === 'inside' ? '근무지 내' : '근무지 외',
    };
  } catch (problem) {
    invalidateResult('입력 조건을 확인하세요.');
    error.textContent = problem.message;
    error.hidden = false;
  }
});

form.addEventListener('reset', () => {
  setTimeout(() => {
    lodgingRows.replaceChildren();
    addLodging();
    for (const details of form.querySelectorAll('details')) details.open = false;
    setTripType();
    updateDuration();
    invalidateResult('출장 일정과 조건을 입력하세요.');
  });
});

get('open-criteria').addEventListener('click', () => get('criteria-dialog').showModal());
get('add-lodging').addEventListener('click', () => addLodging().querySelector('.rate').focus());
get('print-result').addEventListener('click', () => window.print());
get('save-xlsx').addEventListener('click', async () => {
  const result = currentResult;
  if (!result) return;
  const button = get('save-xlsx');
  button.disabled = true;
  get('export-error').hidden = true;
  try {
    await downloadTripWorkbook(result);
  } catch {
    if (currentResult === result) get('export-error').hidden = false;
  } finally {
    if (currentResult === result) button.disabled = false;
  }
});
addLodging();
setTripType();
updateDuration();
