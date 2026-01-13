// Основной скрипт приложения - инициализация и обработчики событий

// Глобальные переменные и состояния
let currentCalculationType = 'plant';
let calculationResults = null;

// Глобальные элементы DOM (объявляем как переменные для глобального доступа)
let projectNameInput, fullNameInput, typeButtons, calculationSections;
let resultsSection, projectSummary, keyResults, detailedResults, parametersList;
let yearsSliders, yearsDisplays;
let calculateButtons;
let exportJsonBtn, exportDocxBtn, printResultsBtn, newCalculationBtn;
let formulaModal, formulaContent, closeModalBtn;

// Функции для показа формул (глобальные для доступа из HTML)
window.showFormula = function(title, formula) {
    if (!formulaContent) return;
    formulaContent.innerHTML = `
        <h4>${title}</h4>
        <div style="background: #f1f8e9; padding: 15px; border-radius: 5px; font-family: monospace; margin: 10px 0;">
            ${formula}
        </div>
    `;
    if (formulaModal) formulaModal.classList.remove('hidden');
};

window.showPlantFormulas = function() {
    if (!formulaContent) return;
    formulaContent.innerHTML = `
        <h4>Формулы для растениеводства:</h4>
        <div style="margin-top: 15px;">
            <p><strong>Площадь в год N:</strong> Площадь_1 + (Прирост × (N-1))</p>
            <p><strong>Валовой сбор (т):</strong> Площадь × Урожайность / 10</p>
            <p><strong>Товарная продукция (т):</strong> Валовой сбор × Коэффициент товарности</p>
            <p><strong>Выручка (тыс. руб.):</strong> Товарная продукция × Цена / 1000</p>
            <p><strong>Прибыль до налогообложения:</strong> Выручка - Все расходы</p>
            <p><strong>Налоги:</strong> Прибыль × Налоговая ставка</p>
            <p><strong>Чистая прибыль:</strong> Прибыль - Налоги</p>
        </div>
    `;
    if (formulaModal) formulaModal.classList.remove('hidden');
};

window.showAnimalFormulas = function() {
    if (!formulaContent) return;
    formulaContent.innerHTML = `
        <h4>Формулы для животноводства:</h4>
        <div style="margin-top: 15px;">
            <p><strong>Поголовье в год N:</strong> Поголовье_1 + (Прирост × (N-1))</p>
            <p><strong>Производство молока (т):</strong> Поголовье × Период лактации × Надой / 1000</p>
            <p><strong>Выручка от молока:</strong> Молоко × Цена молока / 1000</p>
            <p><strong>Производство мяса (т):</strong> Поголовье × Привес / 1000</p>
            <p><strong>Выручка от мяса:</strong> Мясо × Цена мяса / 1000</p>
            <p><strong>Общая выручка:</strong> Выручка_молоко + Выручка_мясо</p>
            <p><strong>Прибыль до налогообложения:</strong> Выручка - Все расходы</p>
        </div>
    `;
    if (formulaModal) formulaModal.classList.remove('hidden');
};

window.showProductionFormulas = function() {
    if (!formulaContent) return;
    formulaContent.innerHTML = `
        <h4>Формулы для переработки:</h4>
        <div style="margin-top: 15px;">
            <p><strong>Мощность в год N:</strong> Мощность_1 + (Рост мощности × (N-1))</p>
            <p><strong>Выпуск продукции (т):</strong> Мощность × Рабочих дней</p>
            <p><strong>Выручка (тыс. руб.):</strong> Выпуск продукции × Цена / 1000</p>
            <p><strong>Расходы:</strong> Сырье + ФОТ + Прочие расходы</p>
            <p><strong>Прибыль до налогообложения:</strong> Выручка - Все расходы</p>
            <p><strong>Налоги:</strong> Прибыль × Налоговая ставка</p>
            <p><strong>Чистая прибыль:</strong> Прибыль - Налоги</p>
        </div>
    `;
    if (formulaModal) formulaModal.classList.remove('hidden');
};

document.addEventListener('DOMContentLoaded', function() {
    // Инициализация элементов DOM
    projectNameInput = document.getElementById('projectName');
    fullNameInput = document.getElementById('fullName');
    typeButtons = document.querySelectorAll('.type-btn');
    calculationSections = document.querySelectorAll('.calculation-section');
    resultsSection = document.getElementById('resultsSection');
    projectSummary = document.getElementById('projectSummary');
    keyResults = document.getElementById('keyResults');
    detailedResults = document.getElementById('detailedResults');
    parametersList = document.getElementById('parametersList');

    // Элементы управления годами
    yearsSliders = {
        plant: document.getElementById('yearsCountPlant'),
        animal: document.getElementById('yearsCountAnimal'),
        production: document.getElementById('yearsCountProduction')
    };

    yearsDisplays = {
        plant: document.getElementById('yearsDisplayPlant'),
        animal: document.getElementById('yearsDisplayAnimal'),
        production: document.getElementById('yearsDisplayProduction')
    };

    // Кнопки расчета
    calculateButtons = {
        plant: document.getElementById('calculatePlant'),
        animal: document.getElementById('calculateAnimal'),
        production: document.getElementById('calculateProduction')
    };

    // Кнопки действий
    exportJsonBtn = document.getElementById('exportJson');
    exportDocxBtn = document.getElementById('exportDocx');
    printResultsBtn = document.getElementById('printResults');
    newCalculationBtn = document.getElementById('newCalculation');

    // Модальное окно
    formulaModal = document.getElementById('formulaModal');
    formulaContent = document.getElementById('formulaContent');
    closeModalBtn = document.querySelector('.close-modal');

    // Инициализация
    resultsSection.classList.add('hidden');
    formulaModal.classList.add('hidden');

    // Настройка слайдеров лет
    Object.keys(yearsSliders).forEach(type => {
        if (yearsSliders[type]) {
            yearsSliders[type].addEventListener('input', function() {
                if (yearsDisplays[type]) {
                    yearsDisplays[type].textContent = `${this.value} лет`;
                }
            });
        }
    });

    // Переключение типа расчета
    typeButtons.forEach(btn => {
        btn.addEventListener('click', function() {
            typeButtons.forEach(b => b.classList.remove('active'));
            this.classList.add('active');

            calculationSections.forEach(section => section.classList.remove('active'));

            currentCalculationType = this.dataset.type;
            document.getElementById(`${currentCalculationType}-section`).classList.add('active');

            // Скрываем результаты при переключении
            resultsSection.classList.add('hidden');
        });
    });

    // Расчет ФОТ
    function calculateSalary(employees, year) {
        const monthlySalary = 30000 + (year - 1) * 2000; // Индексация 2000 руб. в год
        const annualSalary = monthlySalary * 12 / 1000; // тыс. руб.
        const ndfl = annualSalary * 0.13; // тыс. руб.
        const insurance = annualSalary * 0.3; // тыс. руб.
        const totalPerEmployee = annualSalary + ndfl + insurance;

        return totalPerEmployee * employees;
    }

    // Валидация ввода
    function validateProjectInfo() {
        if (!projectNameInput || !projectNameInput.value.trim()) {
            alert('Пожалуйста, введите название проекта');
            if (projectNameInput) projectNameInput.focus();
            return false;
        }

        if (!fullNameInput || !fullNameInput.value.trim()) {
            alert('Пожалуйста, введите ФИО руководителя');
            if (fullNameInput) fullNameInput.focus();
            return false;
        }

        return true;
    }

    // Расчет растениеводства
    calculateButtons.plant.addEventListener('click', function() {
        if (!validateProjectInfo()) return;

        const yearsCount = parseInt(yearsSliders.plant.value);
        const grantAmount = parseFloat(document.getElementById('grantAmountPlant').value);
        const areaYear1 = parseFloat(document.getElementById('areaYear1').value);
        const yieldPerHa = parseFloat(document.getElementById('yieldPerHa').value);
        const pricePerTon = parseFloat(document.getElementById('pricePerTonPlant').value);
        const areaGrowth = parseFloat(document.getElementById('areaGrowth').value);
        const productCoefficient = parseFloat(document.getElementById('productCoefficientPlant').value);
        const taxRate = parseFloat(document.getElementById('taxRatePlant').value);
        const rawMaterials = parseFloat(document.getElementById('rawMaterialsPlant').value);
        const otherCosts = parseFloat(document.getElementById('otherCostsPlant').value);

        // Собственные средства (10% от общей суммы при условии, что грант 90%)
        const ownFunds = grantAmount / 0.9 * 0.1;
        const totalBudget = grantAmount + ownFunds;

        // Расчет по годам
        const results = [];
        for (let year = 1; year <= yearsCount; year++) {
            const area = areaYear1 + (areaGrowth * (year - 1));
            const grossYield = area * yieldPerHa / 10; // в тоннах
            const marketableProduct = grossYield * productCoefficient;
            const revenue = marketableProduct * pricePerTon / 1000; // в тыс. руб.

            // Расходы
            const salary = calculateSalary(1, year); // ФОТ для 1 работника
            const totalCosts = rawMaterials + salary + otherCosts;

            // Прибыль и налоги
            const profitBeforeTax = revenue - totalCosts;
            const taxes = profitBeforeTax > 0 ? profitBeforeTax * taxRate : 0;
            const netProfit = profitBeforeTax - taxes;

            results.push({
                year,
                area,
                grossYield,
                marketableProduct,
                revenue,
                rawMaterials,
                salary,
                otherCosts,
                totalCosts,
                profitBeforeTax,
                taxes,
                netProfit
            });
        }

        // Итоговые показатели
        const totalRevenue = results.reduce((sum, r) => sum + r.revenue, 0);
        const totalNetProfit = results.reduce((sum, r) => sum + r.netProfit, 0);
        const avgNetProfit = totalNetProfit / yearsCount;
        const avgCosts = results.reduce((sum, r) => sum + r.totalCosts, 0) / yearsCount;
        const profitability = avgCosts > 0 ? (avgNetProfit / avgCosts) * 100 : 0;
        const paybackPeriod = avgNetProfit > 0 ? totalBudget / avgNetProfit : 0;

        calculationResults = {
            type: 'plant',
            typeName: 'Растениеводство',
            yearsCount,
            projectName: projectNameInput.value,
            fullName: fullNameInput.value,
            grantAmount,
            ownFunds,
            totalBudget,
            parameters: {
                areaYear1,
                yieldPerHa,
                pricePerTon,
                areaGrowth,
                productCoefficient,
                taxRate,
                rawMaterials,
                otherCosts
            },
            results,
            totalRevenue,
            totalNetProfit,
            avgNetProfit,
            avgCosts,
            profitability,
            paybackPeriod
        };

        displayResults();
    });

    // Расчет животноводства
    calculateButtons.animal.addEventListener('click', function() {
        if (!validateProjectInfo()) return;

        const yearsCount = parseInt(yearsSliders.animal.value);
        const grantAmount = parseFloat(document.getElementById('grantAmountAnimal').value);
        const livestockYear1 = parseFloat(document.getElementById('livestockYear1').value);
        const milkPerAnimal = parseFloat(document.getElementById('milkPerAnimal').value);
        const weightGain = parseFloat(document.getElementById('weightGain').value);
        const livestockGrowth = parseFloat(document.getElementById('livestockGrowth').value);
        const milkPrice = parseFloat(document.getElementById('milkPrice').value);
        const meatPrice = parseFloat(document.getElementById('meatPrice').value);
        const lactationPeriod = parseFloat(document.getElementById('lactationPeriod').value);
        const productCoefficient = parseFloat(document.getElementById('productCoefficientAnimal').value);
        const taxRate = parseFloat(document.getElementById('taxRateAnimal').value);
        const rawMaterials = parseFloat(document.getElementById('rawMaterialsAnimal').value);
        const otherCosts = parseFloat(document.getElementById('otherCostsAnimal').value);

        // Собственные средства
        const ownFunds = grantAmount / 0.9 * 0.1;
        const totalBudget = grantAmount + ownFunds;

        // Расчет по годам
        const results = [];
        for (let year = 1; year <= yearsCount; year++) {
            const livestock = livestockYear1 + (livestockGrowth * (year - 1));

            // Молоко
            const milkProduction = livestock * lactationPeriod * milkPerAnimal / 1000; // в тоннах
            const marketableMilk = milkProduction * productCoefficient;
            const milkRevenue = marketableMilk * milkPrice / 1000; // в тыс. руб.

            // Мясо
            const meatProduction = livestock * weightGain / 1000; // в тоннах
            const marketableMeat = meatProduction * productCoefficient;
            const meatRevenue = marketableMeat * meatPrice / 1000; // в тыс. руб.

            // Общая выручка
            const revenue = milkRevenue + meatRevenue;

            // Расходы
            const salary = calculateSalary(2, year); // ФОТ для 2 работников
            const totalCosts = rawMaterials + salary + otherCosts;

            // Прибыль и налоги
            const profitBeforeTax = revenue - totalCosts;
            const taxes = profitBeforeTax > 0 ? profitBeforeTax * taxRate : 0;
            const netProfit = profitBeforeTax - taxes;

            results.push({
                year,
                livestock,
                milkProduction,
                marketableMilk,
                milkRevenue,
                meatProduction,
                marketableMeat,
                meatRevenue,
                revenue,
                rawMaterials,
                salary,
                otherCosts,
                totalCosts,
                profitBeforeTax,
                taxes,
                netProfit
            });
        }

        // Итоговые показатели
        const totalRevenue = results.reduce((sum, r) => sum + r.revenue, 0);
        const totalNetProfit = results.reduce((sum, r) => sum + r.netProfit, 0);
        const avgNetProfit = totalNetProfit / yearsCount;
        const avgCosts = results.reduce((sum, r) => sum + r.totalCosts, 0) / yearsCount;
        const profitability = avgCosts > 0 ? (avgNetProfit / avgCosts) * 100 : 0;
        const paybackPeriod = avgNetProfit > 0 ? totalBudget / avgNetProfit : 0;

        calculationResults = {
            type: 'animal',
            typeName: 'Животноводство',
            yearsCount,
            projectName: projectNameInput.value,
            fullName: fullNameInput.value,
            grantAmount,
            ownFunds,
            totalBudget,
            parameters: {
                livestockYear1,
                milkPerAnimal,
                weightGain,
                livestockGrowth,
                milkPrice,
                meatPrice,
                lactationPeriod,
                productCoefficient,
                taxRate,
                rawMaterials,
                otherCosts
            },
            results,
            totalRevenue,
            totalNetProfit,
            avgNetProfit,
            avgCosts,
            profitability,
            paybackPeriod
        };

        displayResults();
    });

    // Расчет переработки
    calculateButtons.production.addEventListener('click', function() {
        if (!validateProjectInfo()) return;

        const yearsCount = parseInt(yearsSliders.production.value);
        const grantAmount = parseFloat(document.getElementById('grantAmountProduction').value);
        const capacityYear1 = parseFloat(document.getElementById('capacityYear1').value);
        const capacityGrowth = parseFloat(document.getElementById('capacityGrowth').value);
        const pricePerTon = parseFloat(document.getElementById('pricePerTonProduction').value);
        const workingDays = parseFloat(document.getElementById('workingDays').value);
        const rawMaterials = parseFloat(document.getElementById('rawMaterialsProduction').value);
        const taxRate = parseFloat(document.getElementById('taxRateProduction').value);
        const otherCosts = parseFloat(document.getElementById('otherCostsProduction').value);

        // Собственные средства
        const ownFunds = grantAmount / 0.9 * 0.1;
        const totalBudget = grantAmount + ownFunds;

        // Расчет по годам
        const results = [];
        for (let year = 1; year <= yearsCount; year++) {
            const capacity = capacityYear1 + (capacityGrowth * (year - 1));
            const production = capacity * workingDays; // в тоннах
            const revenue = production * pricePerTon / 1000; // в тыс. руб.

            // Расходы
            const salary = calculateSalary(2, year); // ФОТ для 2 работников
            const totalCosts = rawMaterials + salary + otherCosts;

            // Прибыль и налоги
            const profitBeforeTax = revenue - totalCosts;
            const taxes = profitBeforeTax > 0 ? profitBeforeTax * taxRate : 0;
            const netProfit = profitBeforeTax - taxes;

            results.push({
                year,
                capacity,
                production,
                revenue,
                rawMaterials,
                salary,
                otherCosts,
                totalCosts,
                profitBeforeTax,
                taxes,
                netProfit
            });
        }

        // Итоговые показатели
        const totalRevenue = results.reduce((sum, r) => sum + r.revenue, 0);
        const totalNetProfit = results.reduce((sum, r) => sum + r.netProfit, 0);
        const avgNetProfit = totalNetProfit / yearsCount;
        const avgCosts = results.reduce((sum, r) => sum + r.totalCosts, 0) / yearsCount;
        const profitability = avgCosts > 0 ? (avgNetProfit / avgCosts) * 100 : 0;
        const paybackPeriod = avgNetProfit > 0 ? totalBudget / avgNetProfit : 0;

        calculationResults = {
            type: 'production',
            typeName: 'Переработка сельскохозяйственной продукции',
            yearsCount,
            projectName: projectNameInput.value,
            fullName: fullNameInput.value,
            grantAmount,
            ownFunds,
            totalBudget,
            parameters: {
                capacityYear1,
                capacityGrowth,
                pricePerTon,
                workingDays,
                rawMaterials,
                taxRate,
                otherCosts
            },
            results,
            totalRevenue,
            totalNetProfit,
            avgNetProfit,
            avgCosts,
            profitability,
            paybackPeriod
        };

        displayResults();
    });

    // Функции для отображения результатов
    function displayResults() {
        if (!calculationResults || !resultsSection || !projectSummary || !keyResults || !detailedResults || !parametersList) return;

        // Отображаем секцию результатов
        resultsSection.classList.remove('hidden');

        // Информация о проекте
        projectSummary.innerHTML = `
            <div class="result-card">
                <div class="result-label">Проект</div>
                <div class="result-value">${calculationResults.projectName}</div>
            </div>
            <div class="result-card">
                <div class="result-label">Руководитель</div>
                <div class="result-value">${calculationResults.fullName}</div>
            </div>
            <div class="result-card">
                <div class="result-label">Направление</div>
                <div class="result-value">${calculationResults.typeName}</div>
            </div>
        `;

        // Ключевые показатели с подсказками
        keyResults.innerHTML = `
            <div class="result-card">
                <div class="result-label">Общий бюджет проекта</div>
                <div class="result-value">${calculationResults.totalBudget.toFixed(0)} тыс. руб.</div>
                <div class="info-icon tooltip" onclick="showFormula('Общий бюджет проекта', 'Грант + Собственные средства = ${calculationResults.grantAmount.toFixed(0)} + ${calculationResults.ownFunds.toFixed(0)} = ${calculationResults.totalBudget.toFixed(0)} тыс. руб.')">
                    <i class="fas fa-info-circle"></i>
                </div>
            </div>
            <div class="result-card">
                <div class="result-label">Средства гранта</div>
                <div class="result-value">${calculationResults.grantAmount.toFixed(0)} тыс. руб.</div>
            </div>
            <div class="result-card">
                <div class="result-label">Собственные средства</div>
                <div class="result-value">${calculationResults.ownFunds.toFixed(0)} тыс. руб.</div>
                <div class="info-icon tooltip" onclick="showFormula('Собственные средства', 'Собственные средства = Грант / 0.9 × 0.1 = ${calculationResults.grantAmount.toFixed(0)} / 0.9 × 0.1 = ${calculationResults.ownFunds.toFixed(0)} тыс. руб.')">
                    <i class="fas fa-info-circle"></i>
                </div>
            </div>
            <div class="result-card">
                <div class="result-label">Среднегодовая чистая прибыль</div>
                <div class="result-value">${calculationResults.avgNetProfit.toFixed(0)} тыс. руб.</div>
                <div class="info-icon tooltip" onclick="showFormula('Среднегодовая чистая прибыль', '∑(Чистая прибыль за все годы) / Количество лет = ${calculationResults.totalNetProfit.toFixed(0)} / ${calculationResults.yearsCount} = ${calculationResults.avgNetProfit.toFixed(0)} тыс. руб.')">
                    <i class="fas fa-info-circle"></i>
                </div>
            </div>
            <div class="result-card">
                <div class="result-label">Рентабельность проекта</div>
                <div class="result-value">${calculationResults.profitability.toFixed(1)}%</div>
                <div class="info-icon tooltip" onclick="showFormula('Рентабельность проекта', 'Среднегодовая чистая прибыль / Среднегодовые затраты × 100% = ${calculationResults.avgNetProfit.toFixed(0)} / ${calculationResults.avgCosts.toFixed(0)} × 100% = ${calculationResults.profitability.toFixed(1)}%')">
                    <i class="fas fa-info-circle"></i>
                </div>
            </div>
            <div class="result-card">
                <div class="result-label">Срок окупаемости</div>
                <div class="result-value">${calculationResults.paybackPeriod.toFixed(1)} лет</div>
                <div class="info-icon tooltip" onclick="showFormula('Срок окупаемости', 'Общий бюджет проекта / Среднегодовая чистая прибыль = ${calculationResults.totalBudget.toFixed(0)} / ${calculationResults.avgNetProfit.toFixed(0)} = ${calculationResults.paybackPeriod.toFixed(1)} лет')">
                    <i class="fas fa-info-circle"></i>
                </div>
            </div>
        `;

        // Детализированные результаты по годам
        let tableHtml = '';
        const years = Array.from({length: calculationResults.yearsCount}, (_, i) => i + 1);

        if (calculationResults.type === 'plant') {
            tableHtml = `
                <table>
                    <thead>
                        <tr>
                            <th>Показатель <span class="info-icon tooltip" onclick="showPlantFormulas()"><i class="fas fa-info-circle"></i></span></th>
                            ${years.map(year => `<th>${year} год</th>`).join('')}
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>Площадь, га</td>
                            ${calculationResults.results.map(r => `<td>${r.area.toFixed(1)}</td>`).join('')}
                        </tr>
                        <tr>
                            <td>Валовой сбор, т</td>
                            ${calculationResults.results.map(r => `<td>${r.grossYield.toFixed(1)}</td>`).join('')}
                        </tr>
                        <tr>
                            <td>Выручка, тыс. руб.</td>
                            ${calculationResults.results.map(r => `<td>${r.revenue.toFixed(0)}</td>`).join('')}
                        </tr>
                        <tr>
                            <td>Расходы, тыс. руб.</td>
                            ${calculationResults.results.map(r => `<td>${r.totalCosts.toFixed(0)}</td>`).join('')}
                        </tr>
                        <tr>
                            <td>Прибыль до налогообложения, тыс. руб.</td>
                            ${calculationResults.results.map(r => `<td>${r.profitBeforeTax.toFixed(0)}</td>`).join('')}
                        </tr>
                        <tr>
                            <td>Налоги, тыс. руб.</td>
                            ${calculationResults.results.map(r => `<td>${r.taxes.toFixed(0)}</td>`).join('')}
                        </tr>
                        <tr class="total-row">
                            <td>Чистая прибыль, тыс. руб.</td>
                            ${calculationResults.results.map(r => `<td>${r.netProfit.toFixed(0)}</td>`).join('')}
                        </tr>
                    </tbody>
                </table>
            `;
        } else if (calculationResults.type === 'animal') {
            tableHtml = `
                <table>
                    <thead>
                        <tr>
                            <th>Показатель <span class="info-icon tooltip" onclick="showAnimalFormulas()"><i class="fas fa-info-circle"></i></span></th>
                            ${years.map(year => `<th>${year} год</th>`).join('')}
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>Поголовье, гол.</td>
                            ${calculationResults.results.map(r => `<td>${r.livestock.toFixed(0)}</td>`).join('')}
                        </tr>
                        <tr>
                            <td>Производство молока, т</td>
                            ${calculationResults.results.map(r => `<td>${r.milkProduction.toFixed(1)}</td>`).join('')}
                        </tr>
                        <tr>
                            <td>Выручка от молока, тыс. руб.</td>
                            ${calculationResults.results.map(r => `<td>${r.milkRevenue.toFixed(0)}</td>`).join('')}
                        </tr>
                        <tr>
                            <td>Производство мяса, т</td>
                            ${calculationResults.results.map(r => `<td>${r.meatProduction.toFixed(1)}</td>`).join('')}
                        </tr>
                        <tr>
                            <td>Выручка от мяса, тыс. руб.</td>
                            ${calculationResults.results.map(r => `<td>${r.meatRevenue.toFixed(0)}</td>`).join('')}
                        </tr>
                        <tr>
                            <td>Общая выручка, тыс. руб.</td>
                            ${calculationResults.results.map(r => `<td>${r.revenue.toFixed(0)}</td>`).join('')}
                        </tr>
                        <tr>
                            <td>Расходы, тыс. руб.</td>
                            ${calculationResults.results.map(r => `<td>${r.totalCosts.toFixed(0)}</td>`).join('')}
                        </tr>
                        <tr class="total-row">
                            <td>Чистая прибыль, тыс. руб.</td>
                            ${calculationResults.results.map(r => `<td>${r.netProfit.toFixed(0)}</td>`).join('')}
                        </tr>
                    </tbody>
                </table>
            `;
        } else if (calculationResults.type === 'production') {
            tableHtml = `
                <table>
                    <thead>
                        <tr>
                            <th>Показатель <span class="info-icon tooltip" onclick="showProductionFormulas()"><i class="fas fa-info-circle"></i></span></th>
                            ${years.map(year => `<th>${year} год</th>`).join('')}
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>Производственная мощность, т/сут.</td>
                            ${calculationResults.results.map(r => `<td>${r.capacity.toFixed(4)}</td>`).join('')}
                        </tr>
                        <tr>
                            <td>Выпуск продукции, т</td>
                            ${calculationResults.results.map(r => `<td>${r.production.toFixed(1)}</td>`).join('')}
                        </tr>
                        <tr>
                            <td>Выручка, тыс. руб.</td>
                            ${calculationResults.results.map(r => `<td>${r.revenue.toFixed(0)}</td>`).join('')}
                        </tr>
                        <tr>
                            <td>Расходы, тыс. руб.</td>
                            ${calculationResults.results.map(r => `<td>${r.totalCosts.toFixed(0)}</td>`).join('')}
                        </tr>
                        <tr>
                            <td>Прибыль до налогообложения, тыс. руб.</td>
                            ${calculationResults.results.map(r => `<td>${r.profitBeforeTax.toFixed(0)}</td>`).join('')}
                        </tr>
                        <tr>
                            <td>Налоги, тыс. руб.</td>
                            ${calculationResults.results.map(r => `<td>${r.taxes.toFixed(0)}</td>`).join('')}
                        </tr>
                        <tr class="total-row">
                            <td>Чистая прибыль, тыс. руб.</td>
                            ${calculationResults.results.map(r => `<td>${r.netProfit.toFixed(0)}</td>`).join('')}
                        </tr>
                    </tbody>
                </table>
            `;
        }

        detailedResults.innerHTML = tableHtml;

        // Вспомогательные функции для отображения параметров
        function getParameterName(key) {
            const names = {
                areaYear1: 'Площадь в 1-й год (га)',
                yieldPerHa: 'Урожайность (ц/га)',
                pricePerTon: 'Цена за тонну (руб.)',
                areaGrowth: 'Прирост площади (га/год)',
                productCoefficient: 'Коэффициент товарности',
                taxRate: 'Налоговая ставка',
                rawMaterials: 'Сырье и материалы (тыс. руб./год)',
                otherCosts: 'Прочие расходы (тыс. руб./год)',
                livestockYear1: 'Поголовье в 1-й год (гол.)',
                milkPerAnimal: 'Надой на голову (кг/день)',
                weightGain: 'Привес на голову (кг/год)',
                livestockGrowth: 'Прирост поголовья (гол./год)',
                milkPrice: 'Цена молока (руб./т)',
                meatPrice: 'Цена мяса (руб./т)',
                lactationPeriod: 'Период лактации (дней)',
                capacityYear1: 'Мощность в 1-й год (т/сут.)',
                capacityGrowth: 'Рост мощности (т/сут./год)',
                workingDays: 'Рабочих дней в году'
            };
            return names[key] || key;
        }

        function formatParameterValue(key, value) {
            if (key.includes('Price') || key.includes('price')) {
                return `${value.toLocaleString('ru-RU')}`;
            } else if (key.includes('Rate') || key === 'productCoefficient') {
                return `${(value * 100).toFixed(1)}%`;
            } else if (key.includes('Growth') || key.includes('area') || key.includes('yield') ||
                      key.includes('capacity') || key.includes('gain') || key.includes('Period')) {
                return value;
            } else if (typeof value === 'number') {
                return `${value.toLocaleString('ru-RU')}`;
            }
            return value;
        }

        // Список параметров
        parametersList.innerHTML = `
            <h4><i class="fas fa-cog"></i> Использованные параметры расчета:</h4>
            ${Object.entries(calculationResults.parameters).map(([key, value]) => `
                <div class="parameter-item">
                    <span class="parameter-name">${getParameterName(key)}:</span>
                    <span class="parameter-value">${formatParameterValue(key, value)}</span>
                </div>
            `).join('')}
        `;

        // Прокрутка к результатам
        resultsSection.scrollIntoView({ behavior: 'smooth' });
    }

    // Экспорт в JSON
    exportJsonBtn.addEventListener('click', function() {
        if (!calculationResults) {
            alert('Сначала выполните расчет');
            return;
        }

        const dataStr = JSON.stringify(calculationResults, null, 2);
        const dataBlob = new Blob([dataStr], {type: 'application/json'});

        saveAs(dataBlob, `агростартап_${calculationResults.projectName.replace(/\s+/g, '_')}.json`);

        alert('Данные успешно экспортированы в формате JSON');
    });

    // Экспорт в DOCX
    exportDocxBtn.addEventListener('click', async function() {
        if (!calculationResults) {
            alert('Сначала выполните расчет');
            return;
        }

        try {
            await generateDocxReport();
        } catch (error) {
            console.error('Ошибка при создании DOCX:', error);
            alert('Произошла ошибка при создании документа. Проверьте консоль для подробностей.');
        }
    });

    // Функция генерации DOCX
    async function generateDocxReport() {
        const results = calculationResults;

        // Создаем новый документ
        const doc = new docx.Document({
            sections: [{
                properties: {},
                children: []
            }]
        });

        // Заголовок
        const title = new docx.Paragraph({
            text: `Бизнес-план агростартапа: ${results.projectName}`,
            heading: docx.HeadingLevel.HEADING_1,
            alignment: docx.AlignmentType.CENTER,
            spacing: { after: 200 }
        });

        doc.addSection({
            children: [title]
        });

        // Информация о проекте
        const info1 = new docx.Paragraph({
            text: `Руководитель проекта: ${results.fullName}`,
            spacing: { after: 100 }
        });

        const info2 = new docx.Paragraph({
            text: `Направление деятельности: ${results.typeName}`,
            spacing: { after: 100 }
        });

        const info3 = new docx.Paragraph({
            text: `Период расчета: ${results.yearsCount} лет`,
            spacing: { after: 200 }
        });

        doc.addSection({
            children: [info1, info2, info3]
        });

        // Финансирование проекта
        const financeTitle = new docx.Paragraph({
            text: '1. Финансирование проекта',
            heading: docx.HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 100 }
        });

        doc.addSection({
            children: [financeTitle]
        });

        // Таблица финансирования
        const financeTable = new docx.Table({
            rows: [
                new docx.TableRow({
                    children: [
                        new docx.TableCell({
                            children: [new docx.Paragraph('Показатель')],
                            shading: { fill: "4CAF50" }
                        }),
                        new docx.TableCell({
                            children: [new docx.Paragraph('Значение, тыс. руб.')],
                            shading: { fill: "4CAF50" }
                        })
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({
                            children: [new docx.Paragraph('Общий бюджет проекта')]
                        }),
                        new docx.TableCell({
                            children: [new docx.Paragraph(results.totalBudget.toFixed(0))]
                        })
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({
                            children: [new docx.Paragraph('Средства гранта')]
                        }),
                        new docx.TableCell({
                            children: [new docx.Paragraph(results.grantAmount.toFixed(0))]
                        })
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({
                            children: [new docx.Paragraph('Собственные средства')]
                        }),
                        new docx.TableCell({
                            children: [new docx.Paragraph(results.ownFunds.toFixed(0))]
                        })
                    ]
                })
            ],
            width: { size: 100, type: docx.WidthType.PERCENTAGE }
        });

        doc.addSection({
            children: [financeTable]
        });

        // Ключевые показатели
        const indicatorsTitle = new docx.Paragraph({
            text: '2. Ключевые показатели эффективности',
            heading: docx.HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 100 }
        });

        doc.addSection({
            children: [indicatorsTitle]
        });

        // Таблица ключевых показателей
        const indicatorsTable = new docx.Table({
            rows: [
                new docx.TableRow({
                    children: [
                        new docx.TableCell({
                            children: [new docx.Paragraph('Показатель')],
                            shading: { fill: "4CAF50" }
                        }),
                        new docx.TableCell({
                            children: [new docx.Paragraph('Значение')],
                            shading: { fill: "4CAF50" }
                        })
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({
                            children: [new docx.Paragraph('Общая выручка за период')]
                        }),
                        new docx.TableCell({
                            children: [new docx.Paragraph(`${results.totalRevenue.toFixed(0)} тыс. руб.`)]
                        })
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({
                            children: [new docx.Paragraph('Среднегодовая чистая прибыль')]
                        }),
                        new docx.TableCell({
                            children: [new docx.Paragraph(`${results.avgNetProfit.toFixed(0)} тыс. руб.`)]
                        })
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({
                            children: [new docx.Paragraph('Рентабельность проекта')]
                        }),
                        new docx.TableCell({
                            children: [new docx.Paragraph(`${results.profitability.toFixed(1)}%`)]
                        })
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({
                            children: [new docx.Paragraph('Срок окупаемости')]
                        }),
                        new docx.TableCell({
                            children: [new docx.Paragraph(`${results.paybackPeriod.toFixed(1)} лет`)]
                        })
                    ]
                })
            ],
            width: { size: 100, type: docx.WidthType.PERCENTAGE }
        });

        doc.addSection({
            children: [indicatorsTable]
        });

        // Детализированные результаты
        const detailedTitle = new docx.Paragraph({
            text: '3. Детализированные результаты по годам',
            heading: docx.HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 100 }
        });

        doc.addSection({
            children: [detailedTitle]
        });

        // Создаем таблицу с результатами по годам
        const yearHeaders = results.results.map(r =>
            new docx.TableCell({
                children: [new docx.Paragraph(`${r.year} год`)]
            })
        );

        const headerRow = new docx.TableRow({
            children: [
                new docx.TableCell({
                    children: [new docx.Paragraph('Показатель')],
                    shading: { fill: "4CAF50" }
                }),
                ...yearHeaders
            ]
        });

        let dataRows = [];

        if (results.type === 'plant') {
            dataRows = [
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ children: [new docx.Paragraph('Площадь, га')] }),
                        ...results.results.map(r =>
                            new docx.TableCell({ children: [new docx.Paragraph(r.area.toFixed(1))] })
                        )
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ children: [new docx.Paragraph('Валовой сбор, т')] }),
                        ...results.results.map(r =>
                            new docx.TableCell({ children: [new docx.Paragraph(r.grossYield.toFixed(1))] })
                        )
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ children: [new docx.Paragraph('Выручка, тыс. руб.')] }),
                        ...results.results.map(r =>
                            new docx.TableCell({ children: [new docx.Paragraph(r.revenue.toFixed(0))] })
                        )
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ children: [new docx.Paragraph('Чистая прибыль, тыс. руб.')] }),
                        ...results.results.map(r =>
                            new docx.TableCell({ children: [new docx.Paragraph(r.netProfit.toFixed(0))] })
                        )
                    ]
                })
            ];
        } else if (results.type === 'animal') {
            dataRows = [
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ children: [new docx.Paragraph('Поголовье, гол.')] }),
                        ...results.results.map(r =>
                            new docx.TableCell({ children: [new docx.Paragraph(r.livestock.toFixed(0))] })
                        )
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ children: [new docx.Paragraph('Производство молока, т')] }),
                        ...results.results.map(r =>
                            new docx.TableCell({ children: [new docx.Paragraph(r.milkProduction.toFixed(1))] })
                        )
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ children: [new docx.Paragraph('Общая выручка, тыс. руб.')] }),
                        ...results.results.map(r =>
                            new docx.TableCell({ children: [new docx.Paragraph(r.revenue.toFixed(0))] })
                        )
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ children: [new docx.Paragraph('Чистая прибыль, тыс. руб.')] }),
                        ...results.results.map(r =>
                            new docx.TableCell({ children: [new docx.Paragraph(r.netProfit.toFixed(0))] })
                        )
                    ]
                })
            ];
        } else if (results.type === 'production') {
            dataRows = [
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ children: [new docx.Paragraph('Производственная мощность, т/сут.')] }),
                        ...results.results.map(r =>
                            new docx.TableCell({ children: [new docx.Paragraph(r.capacity.toFixed(4))] })
                        )
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ children: [new docx.Paragraph('Выпуск продукции, т')] }),
                        ...results.results.map(r =>
                            new docx.TableCell({ children: [new docx.Paragraph(r.production.toFixed(1))] })
                        )
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ children: [new docx.Paragraph('Выручка, тыс. руб.')] }),
                        ...results.results.map(r =>
                            new docx.TableCell({ children: [new docx.Paragraph(r.revenue.toFixed(0))] })
                        )
                    ]
                }),
                new docx.TableRow({
                    children: [
                        new docx.TableCell({ children: [new docx.Paragraph('Чистая прибыль, тыс. руб.')] }),
                        ...results.results.map(r =>
                            new docx.TableCell({ children: [new docx.Paragraph(r.netProfit.toFixed(0))] })
                        )
                    ]
                })
            ];
        }

        const detailedTable = new docx.Table({
            rows: [headerRow, ...dataRows],
            width: { size: 100, type: docx.WidthType.PERCENTAGE }
        });

        doc.addSection({
            children: [detailedTable]
        });

        // Примечание
        const note = new docx.Paragraph({
            text: 'Примечание: Расчеты выполнены с использованием калькулятора агростартапа. Все значения приведены в тыс. рублей, если не указано иное.',
            spacing: { before: 200, after: 100 }
        });

        const date = new docx.Paragraph({
            text: `Дата формирования отчета: ${new Date().toLocaleDateString('ru-RU')}`,
            alignment: docx.AlignmentType.RIGHT
        });

        doc.addSection({
            children: [note, date]
        });

        // Генерируем документ
        const blob = await docx.Packer.toBlob(doc);
        saveAs(blob, `Бизнес_план_${results.projectName.replace(/\s+/g, '_')}.docx`);

        alert('Документ успешно создан и загружен!');
    }

    // Печать результатов
    printResultsBtn.addEventListener('click', function() {
        if (!calculationResults) {
            alert('Сначала выполните расчет');
            return;
        }

        window.print();
    });

    // Новый расчет
    newCalculationBtn.addEventListener('click', function() {
        // Скрываем результаты
        resultsSection.classList.add('hidden');

        // Сбрасываем форму
        calculationResults = null;

        // Прокрутка к началу
        document.querySelector('.app-container').scrollIntoView({ behavior: 'smooth' });
    });

    // Закрытие модального окна
    if (closeModalBtn) {
        closeModalBtn.addEventListener('click', function() {
            formulaModal.classList.add('hidden');
        });
    }

    // Закрытие модального окна при клике вне его
    document.addEventListener('click', function(event) {
        if (formulaModal &&
            !formulaModal.classList.contains('hidden') &&
            event.target === formulaModal) {
            formulaModal.classList.add('hidden');
        }
    });

    // Закрытие модального окна по клавише ESC
    document.addEventListener('keydown', function(event) {
        if (event.key === 'Escape' &&
            formulaModal &&
            !formulaModal.classList.contains('hidden')) {
            formulaModal.classList.add('hidden');
        }
    });
});
