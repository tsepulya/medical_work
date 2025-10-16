import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import glob
import os
import io
from datetime import datetime

# Настройка страницы
st.set_page_config(
    page_title="Обработчик отчетов ИППСУ",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Обработчик отчетов ИППСУ")
st.markdown("Загрузите Excel файлы для автоматической обработки")

# Функции из вашего кода
def create_mini_df(df, start, end):
    """Создание мини-датафрейма из основного"""
    mini_df = df.loc[start:end].copy()
    first_index = mini_df.index[0]
    
    # Устанавливаем заголовки и удаляем первую строку
    mini_df.columns = mini_df.iloc[0]
    mini_df = mini_df.drop(first_index).reset_index(drop=True)
    
    # Удаляем полностью пустые столбцы
    mini_df = mini_df.dropna(axis=1, how='all')
    
    return mini_df

def create_num_names_services(mini_df):
    """Создание справочника услуг"""
    if 'Наименование услуги' not in mini_df.columns or '№ усл.' not in mini_df.columns:
        return pd.DataFrame()
    
    services_df = mini_df[['Наименование услуги', '№ усл.']].copy()
    services_df = services_df.dropna()
    services_df = services_df[services_df['Наименование услуги'] != "Итого по услуге:"]
    services_df = services_df.drop_duplicates()
    
    return services_df

def is_column_empty(df, column_name):
    """Проверка пустоты столбца"""
    if df.empty or column_name not in df.columns:
        return True
    
    for value in df[column_name]:
        if pd.isna(value):
            continue
        if isinstance(value, str) and value.strip().lower() in ['nan', 'null', '']:
            continue
        return False
    
    return True

def process_excel_files(uploaded_files):
    """Основная функция обработки файлов"""
    # Инициализация результирующего датафрейма
    result = pd.DataFrame(columns=[
        'ФИО ребенка', 
        '№ ИППСУ', 
        'Наименование услуги', 
        'Дата оказания', 
        'Кол-во', 
        'Должность специалиста', 
        'Специалист'
    ])
    
    total_files = len(uploaded_files)
    
    for file_idx, uploaded_file in enumerate(uploaded_files):
        try:
            # Создаем прогресс бар для текущего файла
            progress_text = f"Обработка файла {file_idx + 1}/{total_files}: {uploaded_file.name}"
            progress_bar = st.progress(0, text=progress_text)
            
            # Чтение файла
            df = pd.read_excel(uploaded_file)
            
            # Обновляем прогресс
            progress_bar.progress(0.2, text=f"{progress_text} - Чтение файла")
            
            # Предобработка данных
            df.iloc[:, 0] = df.iloc[:, 0].astype(str).str.replace('№ п/п', '№ усл.', regex=False)
            df = df.dropna(axis=1, how='all')
            
            # Извлекаем ФИО ребенка
            child_name = df.iloc[2, 0] if len(df) > 2 else "Неизвестно"
            
            # Поиск разделов "Предоставленные"
            find_text = df[df.iloc[:, 0].str.contains('Предоставленные', na=False)]
            
            progress_bar.progress(0.4, text=f"{progress_text} - Поиск разделов")
            
            # Обработка каждого раздела
            sections_processed = 0
            total_sections = len(find_text.index)
            
            for i in range(len(find_text.index)):
                try:
                    order_num_df = i
                    end_df = len(df) - 2
                    
                    if df.iloc[find_text.index[i], 0] != 'Предоставленные срочные услуги':
                        start_index = find_text.index[order_num_df] + 4
                        end_index = find_text.index[order_num_df + 1] - 3 if order_num_df + 1 < len(find_text.index) else end_df
                        numIPPSU = df.iloc[find_text.index[order_num_df] + 1, 2]
                    else:
                        start_index = find_text.index[order_num_df] + 2
                        end_index = end_df
                        numIPPSU = 'срочные услуги'
                    
                    # Создаем мини-датафрейм
                    test = create_mini_df(df, start_index, end_index)
                    
                    if not is_column_empty(test, "№ усл."):
                        # Создаем справочник услуг
                        services = create_num_names_services(test)
                        
                        if not services.empty:
                            mapping_dict = services.set_index('№ усл.')['Наименование услуги'].to_dict()
                            
                            # Заполняем пустые ячейки
                            test['Наименование услуги'] = test['Наименование услуги'].fillna(
                                test['№ усл.'].map(mapping_dict)
                            )
                            
                            # Удаляем строки с NaN в дате оказания
                            test = test.dropna(subset=['Дата оказания'])
                            
                            # Создаем новые данные
                            new_data = pd.DataFrame({
                                'ФИО ребенка': [child_name] * len(test),
                                '№ ИППСУ': [numIPPSU] * len(test),
                                'Наименование услуги': test['Наименование услуги'],
                                'Дата оказания': test['Дата оказания'],
                                'Кол-во': test['Кол-во'],
                                'Должность специалиста': test['Должность специалиста'],
                                'Специалист': test['Специалист']
                            })
                            
                            # Добавляем к результату
                            result = pd.concat([result, new_data], ignore_index=True)
                            
                            sections_processed += 1
                    
                except Exception as e:
                    st.warning(f"Ошибка в разделе {i+1} файла {uploaded_file.name}: {str(e)}")
                
                # Обновляем прогресс внутри файла
                section_progress = 0.4 + (0.4 * (i + 1) / total_sections) if total_sections > 0 else 0.8
                progress_bar.progress(section_progress, text=f"{progress_text} - Раздел {i+1}/{total_sections}")
            
            progress_bar.progress(1.0, text=f"{progress_text} - Завершено")
            st.success(f"✅ {uploaded_file.name} - обработан ({sections_processed} разделов)")
            
        except Exception as e:
            st.error(f"❌ Ошибка при обработке {uploaded_file.name}: {str(e)}")
    
    return result

# Интерфейс приложения
with st.sidebar:
    st.header("📋 Инструкция")
    st.markdown("""
    1. **Загрузите** Excel файлы отчетов ИППСУ
    2. **Нажмите** кнопку 'Обработать файлы'
    3. **Скачайте** готовый результат
    
    **Поддерживаемые форматы:**
    - .xlsx
    - .xls
    """)
    
    st.header("ℹ️ О приложении")
    st.markdown("""
    Автоматически обрабатывает файлы отчетов:
    - Извлекает данные об услугах
    - Объединяет в единый формат
    - Сохраняет структуру данных
    """)

# Основная область
uploaded_files = st.file_uploader(
    "Выберите файлы отчетов ИППСУ",
    type=['xlsx', 'xls'],
    accept_multiple_files=True,
    help="Можно выбрать несколько файлов одновременно"
)

if uploaded_files:
    st.success(f"📁 Загружено файлов: {len(uploaded_files)}")
    
    # Показываем список файлов
    with st.expander("📋 Список загруженных файлов"):
        for i, file in enumerate(uploaded_files):
            file_size = file.size / 1024  # размер в KB
            st.write(f"{i+1}. **{file.name}** ({file_size:.1f} KB)")
    
    # Кнопка обработки
    if st.button("🚀 Обработать файлы", type="primary", use_container_width=True):
        with st.spinner("Обрабатываю файлы... Это может занять несколько минут"):
            try:
                # Обработка файлов
                result_df = process_excel_files(uploaded_files)
                
                if not result_df.empty:
                    st.balloons()
                    st.success(f"🎉 Обработка завершена! Обработано строк: **{len(result_df):,}**")
                    
                    # Показываем статистику
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Всего строк", f"{len(result_df):,}")
                    with col2:
                        st.metric("Уникальных детей", result_df['ФИО ребенка'].nunique())
                    with col3:
                        st.metric("Услуг", result_df['Наименование услуги'].nunique())
                    with col4:
                        st.metric("Специалистов", result_df['Специалист'].nunique())
                    
                    # Предпросмотр данных
                    st.subheader("👀 Предпросмотр данных")
                    st.dataframe(result_df.head(20), use_container_width=True)
                    
                    # Детальная статистика
                    with st.expander("📈 Детальная статистика"):
                        tab1, tab2, tab3 = st.tabs(["Дети", "Услуги", "Специалисты"])
                        
                        with tab1:
                            st.write("**По детям:**")
                            child_stats = result_df['ФИО ребенка'].value_counts()
                            st.dataframe(child_stats)
                        
                        with tab2:
                            st.write("**По услугам:**")
                            service_stats = result_df['Наименование услуги'].value_counts()
                            st.dataframe(service_stats)
                        
                        with tab3:
                            st.write("**По специалистам:**")
                            specialist_stats = result_df['Специалист'].value_counts()
                            st.dataframe(specialist_stats)
                    
                    # Скачивание результата
                    st.subheader("💾 Скачать результат")
                    
                    # Подготовка файла
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        result_df.to_excel(writer, index=False, sheet_name='Обработанные_данные')
                    output.seek(0)
                    
                    # Кнопки скачивания
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.download_button(
                            label="📥 Скачать как Excel",
                            data=output,
                            file_name=f"обработанные_отчеты_ИППСУ_{timestamp}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            type="primary"
                        )
                    
                    with col2:
                        csv_data = result_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                        st.download_button(
                            label="📥 Скачать как CSV",
                            data=csv_data,
                            file_name=f"обработанные_отчеты_ИППСУ_{timestamp}.csv",
                            mime="text/csv",
                            use_container_width=True
                        )
                
                else:
                    st.warning("⚠️ Не удалось извлечь данные из файлов. Проверьте формат файлов.")
                    
            except Exception as e:
                st.error(f"❌ Критическая ошибка при обработке: {str(e)}")

else:
    st.info("👆 Загрузите Excel файлы отчетов ИППСУ для начала обработки")

# Футер
st.markdown("---")

st.caption("Обработчик отчетов ИППСУ | Создано на Streamlit")
