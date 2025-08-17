"""
Performance Benchmarking Test Suite for AutoWord vNext.

This module implements performance benchmarking tests for large document processing,
including extraction, planning, execution, and validation performance metrics.
"""

import os
import time
import psutil
import pytest
import tempfile
import shutil
import json
import statistics
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from unittest.mock import Mock, patch, MagicMock
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

from autoword.vnext.pipeline import VNextPipeline
from autoword.vnext.executor import DocumentExecutor
from autoword.vnext.extractor import DocumentExtractor
from autoword.vnext.planner import DocumentPlanner
from autoword.vnext.validator import DocumentValidator
from autoword.vnext.models import (
    StructureV1, DocumentMetadata, StyleDefinition, ParagraphSkeleton,
    HeadingReference, FieldReference, TableSkeleton, PlanV1,
    FontSpec, ParagraphSpec, LineSpacingMode, StyleType
)
from autoword.core.llm_client import LLMClient


class PerformanceMetrics:
    """Performance metrics collection and analysis."""
    
    def __init__(self):
        self.metrics = {}
        self.start_times = {}
        self.memory_snapshots = {}
        
    def start_measurement(self, operation_name: str):
        """Start measuring performance for an operation."""
        self.start_times[operation_name] = time.time()
        self.memory_snapshots[f"{operation_name}_start"] = self._get_memory_usage()
        
    def end_measurement(self, operation_name: str) -> Dict[str, Any]:
        """End measurement and return metrics."""
        if operation_name not in self.start_times:
            raise ValueError(f"Measurement for {operation_name} was not started")
            
        end_time = time.time()
        end_memory = self._get_memory_usage()
        start_memory = self.memory_snapshots.get(f"{operation_name}_start", 0)
        
        metrics = {
            'duration_seconds': end_time - self.start_times[operation_name],
            'memory_start_mb': start_memory,
            'memory_end_mb': end_memory,
            'memory_peak_mb': self._get_peak_memory_usage(operation_name),
            'memory_delta_mb': end_memory - start_memory,
            'cpu_usage_percent': self._get_cpu_usage(),
            'timestamp': datetime.now().isoformat()
        }
        
        self.metrics[operation_name] = metrics
        return metrics
        
    def _get_memory_usage(self) -> float:
        """Get current memory usage in MB."""
        try:
            process = psutil.Process()
            return process.memory_info().rss / 1024 / 1024
        except Exception:
            return 0.0
            
    def _get_peak_memory_usage(self, operation_name: str) -> float:
        """Get peak memory usage during operation."""
        # In a real implementation, this would track peak memory
        # For now, return current memory as approximation
        return self._get_memory_usage()
        
    def _get_cpu_usage(self) -> float:
        """Get current CPU usage percentage."""
        try:
            return psutil.cpu_percent(interval=0.1)
        except Exception:
            return 0.0
            
    def get_all_metrics(self) -> Dict[str, Dict[str, Any]]:
        """Get all collected metrics."""
        return self.metrics.copy()
        
    def generate_performance_report(self) -> Dict[str, Any]:
        """Generate comprehensive performance report."""
        if not self.metrics:
            return {'error': 'No metrics collected'}
            
        durations = [m['duration_seconds'] for m in self.metrics.values()]
        memory_deltas = [m['memory_delta_mb'] for m in self.metrics.values()]
        
        return {
            'total_operations': len(self.metrics),
            'total_duration_seconds': sum(durations),
            'average_duration_seconds': statistics.mean(durations),
            'median_duration_seconds': statistics.median(durations),
            'max_duration_seconds': max(durations),
            'min_duration_seconds': min(durations),
            'total_memory_delta_mb': sum(memory_deltas),
            'average_memory_delta_mb': statistics.mean(memory_deltas),
            'max_memory_delta_mb': max(memory_deltas),
            'operations_breakdown': self.metrics
        }


class DocumentGenerator:
    """Generate test documents of various sizes for performance testing."""
    
    @staticmethod
    def generate_large_structure(paragraph_count: int = 1000, 
                               heading_count: int = 50,
                               style_count: int = 20,
                               table_count: int = 10,
                               field_count: int = 15) -> StructureV1:
        """Generate a large document structure for performance testing."""
        
        # Generate styles
        styles = []
        for i in range(style_count):
            styles.append(StyleDefinition(
                name=f"Style_{i}",
                type=StyleType.PARAGRAPH,
                font=FontSpec(
                    east_asian="宋体",
                    latin="Times New Roman",
                    size_pt=12 + (i % 6),  # Vary font sizes
                    bold=(i % 3 == 0)
                ),
                paragraph=ParagraphSpec(
                    line_spacing_mode=LineSpacingMode.MULTIPLE,
                    line_spacing_value=1.0 + (i % 3) * 0.5
                )
            ))
            
        # Generate paragraphs
        paragraphs = []
        for i in range(paragraph_count):
            is_heading = i % 20 == 0  # Every 20th paragraph is a heading
            paragraphs.append(ParagraphSkeleton(
                index=i,
                style_name=f"Style_{i % style_count}",
                preview_text=f"This is paragraph {i} with some sample content for testing purposes.",
                is_heading=is_heading,
                heading_level=(i // 20) % 3 + 1 if is_heading else None
            ))
            
        # Generate headings
        headings = []
        heading_index = 0
        for i in range(0, paragraph_count, 20):
            if heading_index < heading_count:
                headings.append(HeadingReference(
                    paragraph_index=i,
                    level=(heading_index // 10) % 3 + 1,
                    text=f"Heading {heading_index + 1}",
                    style_name=f"Style_{heading_index % style_count}"
                ))
                heading_index += 1
                
        # Generate fields
        fields = []
        for i in range(field_count):
            field_type = ["TOC", "REF", "PAGE", "DATE"][i % 4]
            fields.append(FieldReference(
                paragraph_index=i * (paragraph_count // field_count),
                field_type=field_type,
                field_code=f"{field_type} field code {i}",
                result_text=f"{field_type} result {i}"
            ))
            
        # Generate tables
        tables = []
        for i in range(table_count):
            tables.append(TableSkeleton(
                paragraph_index=i * (paragraph_count // table_count),
                rows=3 + (i % 5),
                columns=2 + (i % 4),
                has_header=(i % 2 == 0)
            ))
            
        return StructureV1(
            metadata=DocumentMetadata(
                title=f"Large Test Document ({paragraph_count} paragraphs)",
                author="Performance Test Generator",
                creation_time=datetime.now(),
                modified_time=datetime.now(),
                page_count=paragraph_count // 20,
                paragraph_count=paragraph_count,
                word_count=paragraph_count * 15  # Estimate 15 words per paragraph
            ),
            styles=styles,
            paragraphs=paragraphs,
            headings=headings,
            fields=fields,
            tables=tables
        )
        
    @staticmethod
    def generate_complex_plan(operation_count: int = 50) -> PlanV1:
        """Generate a complex plan with many operations."""
        from autoword.vnext.models import (
            DeleteSectionByHeading, UpdateToc, DeleteToc, SetStyleRule,
            ReassignParagraphsToStyle, ClearDirectFormatting, MatchMode
        )
        
        operations = []
        
        for i in range(operation_count):
            op_type = i % 6
            
            if op_type == 0:
                operations.append(DeleteSectionByHeading(
                    heading_text=f"Section {i}",
                    level=1,
                    match=MatchMode.EXACT,
                    case_sensitive=False
                ))
            elif op_type == 1:
                operations.append(UpdateToc())
            elif op_type == 2:
                operations.append(DeleteToc(mode="all"))
            elif op_type == 3:
                operations.append(SetStyleRule(
                    target_style_name=f"Style_{i % 10}",
                    font=FontSpec(size_pt=12 + (i % 6)),
                    paragraph=ParagraphSpec(line_spacing_value=1.0 + (i % 3) * 0.5)
                ))
            elif op_type == 4:
                operations.append(ReassignParagraphsToStyle(
                    selector={"style_name": f"Style_{i % 5}"},
                    target_style_name=f"Style_{(i + 1) % 5}",
                    clear_direct_formatting=True
                ))
            else:
                operations.append(ClearDirectFormatting(
                    scope="selection",
                    authorization_required=True
                ))
                
        return PlanV1(ops=operations)


class PerformanceBenchmarkSuite:
    """Main performance benchmarking test suite."""
    
    def __init__(self, output_dir: str = "performance_results"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        self.metrics = PerformanceMetrics()
        self.temp_dir = None
        
        # Performance thresholds
        self.thresholds = {
            'extraction_max_seconds': 30.0,
            'planning_max_seconds': 60.0,
            'execution_max_seconds': 120.0,
            'validation_max_seconds': 30.0,
            'memory_max_mb': 500.0,
            'cpu_max_percent': 80.0
        }
        
    def setup_test_environment(self):
        """Setup test environment."""
        self.temp_dir = Path(tempfile.mkdtemp(prefix="perf_test_"))
        
    def cleanup_test_environment(self):
        """Cleanup test environment."""
        if self.temp_dir and self.temp_dir.exists():
            shutil.rmtree(self.temp_dir)
            
    def run_all_benchmarks(self) -> Dict[str, Any]:
        """Run all performance benchmarks."""
        self.setup_test_environment()
        
        try:
            results = {
                'extraction_benchmarks': self.benchmark_extraction_performance(),
                'planning_benchmarks': self.benchmark_planning_performance(),
                'execution_benchmarks': self.benchmark_execution_performance(),
                'validation_benchmarks': self.benchmark_validation_performance(),
                'pipeline_benchmarks': self.benchmark_pipeline_performance(),
                'scalability_benchmarks': self.benchmark_scalability(),
                'concurrency_benchmarks': self.benchmark_concurrency(),
                'memory_benchmarks': self.benchmark_memory_usage(),
                'overall_assessment': self.generate_performance_assessment()
            }
            
            # Save results
            self.save_benchmark_results(results)
            
            return results
            
        finally:
            self.cleanup_test_environment()
            
    def benchmark_extraction_performance(self) -> Dict[str, Any]:
        """Benchmark document extraction performance."""
        results = {
            'test_cases': [],
            'summary': {},
            'threshold_compliance': {}
        }
        
        # Test different document sizes
        test_sizes = [100, 500, 1000, 2000, 5000]
        
        for size in test_sizes:
            test_name = f"extraction_{size}_paragraphs"
            
            try:
                self.metrics.start_measurement(test_name)
                
                # Mock extraction with large document
                with patch('autoword.vnext.extractor.DocumentExtractor') as mock_extractor_class:
                    mock_extractor = Mock()
                    large_structure = DocumentGenerator.generate_large_structure(
                        paragraph_count=size,
                        heading_count=size // 20,
                        style_count=min(20, size // 50)
                    )
                    
                    mock_extractor.extract_structure.return_value = large_structure
                    mock_extractor.extract_inventory.return_value = Mock()
                    mock_extractor.__enter__.return_value = mock_extractor
                    mock_extractor.__exit__.return_value = None
                    mock_extractor_class.return_value = mock_extractor
                    
                    # Simulate extraction work
                    extractor = mock_extractor_class()
                    with extractor:
                        structure = extractor.extract_structure(f"test_{size}.docx")
                        inventory = extractor.extract_inventory(f"test_{size}.docx")
                        
                    # Simulate processing time based on document size
                    time.sleep(size / 10000)  # Simulate work
                    
                metrics = self.metrics.end_measurement(test_name)
                
                results['test_cases'].append({
                    'test_name': test_name,
                    'document_size': size,
                    'duration_seconds': metrics['duration_seconds'],
                    'memory_delta_mb': metrics['memory_delta_mb'],
                    'meets_threshold': metrics['duration_seconds'] <= self.thresholds['extraction_max_seconds'],
                    'paragraphs_per_second': size / metrics['duration_seconds'] if metrics['duration_seconds'] > 0 else 0
                })
                
            except Exception as e:
                results['test_cases'].append({
                    'test_name': test_name,
                    'document_size': size,
                    'error': str(e),
                    'meets_threshold': False
                })
                
        # Generate summary
        successful_tests = [t for t in results['test_cases'] if 'error' not in t]
        if successful_tests:
            durations = [t['duration_seconds'] for t in successful_tests]
            results['summary'] = {
                'total_tests': len(results['test_cases']),
                'successful_tests': len(successful_tests),
                'average_duration': statistics.mean(durations),
                'max_duration': max(durations),
                'min_duration': min(durations),
                'threshold_violations': len([t for t in successful_tests if not t['meets_threshold']])
            }
            
        return results
        
    def benchmark_planning_performance(self) -> Dict[str, Any]:
        """Benchmark plan generation performance."""
        results = {
            'test_cases': [],
            'summary': {},
            'threshold_compliance': {}
        }
        
        # Test different complexity levels
        complexity_levels = [10, 25, 50, 100, 200]
        
        for complexity in complexity_levels:
            test_name = f"planning_{complexity}_operations"
            
            try:
                self.metrics.start_measurement(test_name)
                
                # Mock planning with complex scenarios
                with patch('autoword.vnext.planner.DocumentPlanner') as mock_planner_class:
                    mock_planner = Mock()
                    complex_plan = DocumentGenerator.generate_complex_plan(complexity)
                    
                    mock_planner.generate_plan.return_value = complex_plan
                    mock_planner.validate_plan_schema.return_value = Mock(is_valid=True)
                    mock_planner_class.return_value = mock_planner
                    
                    # Simulate planning work
                    planner = mock_planner_class(Mock())
                    structure = DocumentGenerator.generate_large_structure(complexity * 10)
                    plan = planner.generate_plan(structure, f"Complex modification with {complexity} operations")
                    
                    # Simulate processing time based on complexity
                    time.sleep(complexity / 1000)  # Simulate work
                    
                metrics = self.metrics.end_measurement(test_name)
                
                results['test_cases'].append({
                    'test_name': test_name,
                    'complexity_level': complexity,
                    'duration_seconds': metrics['duration_seconds'],
                    'memory_delta_mb': metrics['memory_delta_mb'],
                    'meets_threshold': metrics['duration_seconds'] <= self.thresholds['planning_max_seconds'],
                    'operations_per_second': complexity / metrics['duration_seconds'] if metrics['duration_seconds'] > 0 else 0
                })
                
            except Exception as e:
                results['test_cases'].append({
                    'test_name': test_name,
                    'complexity_level': complexity,
                    'error': str(e),
                    'meets_threshold': False
                })
                
        return results
        
    def benchmark_execution_performance(self) -> Dict[str, Any]:
        """Benchmark plan execution performance."""
        results = {
            'test_cases': [],
            'summary': {},
            'threshold_compliance': {}
        }
        
        # Test different operation counts
        operation_counts = [5, 15, 30, 60, 100]
        
        for op_count in operation_counts:
            test_name = f"execution_{op_count}_operations"
            
            try:
                self.metrics.start_measurement(test_name)
                
                # Mock execution with multiple operations
                with patch('autoword.vnext.executor.DocumentExecutor') as mock_executor_class:
                    mock_executor = Mock()
                    
                    # Mock successful execution
                    mock_executor.execute_plan.return_value = f"/modified_{op_count}.docx"
                    mock_executor_class.return_value = mock_executor
                    
                    # Simulate execution work
                    executor = mock_executor_class()
                    plan = DocumentGenerator.generate_complex_plan(op_count)
                    result_path = executor.execute_plan(plan, f"test_{op_count}.docx")
                    
                    # Simulate processing time based on operation count
                    time.sleep(op_count / 100)  # Simulate work
                    
                metrics = self.metrics.end_measurement(test_name)
                
                results['test_cases'].append({
                    'test_name': test_name,
                    'operation_count': op_count,
                    'duration_seconds': metrics['duration_seconds'],
                    'memory_delta_mb': metrics['memory_delta_mb'],
                    'meets_threshold': metrics['duration_seconds'] <= self.thresholds['execution_max_seconds'],
                    'operations_per_second': op_count / metrics['duration_seconds'] if metrics['duration_seconds'] > 0 else 0
                })
                
            except Exception as e:
                results['test_cases'].append({
                    'test_name': test_name,
                    'operation_count': op_count,
                    'error': str(e),
                    'meets_threshold': False
                })
                
        return results
        
    def benchmark_validation_performance(self) -> Dict[str, Any]:
        """Benchmark validation performance."""
        results = {
            'test_cases': [],
            'summary': {},
            'threshold_compliance': {}
        }
        
        # Test different document complexities
        complexities = [
            {'paragraphs': 100, 'styles': 10, 'fields': 5},
            {'paragraphs': 500, 'styles': 20, 'fields': 15},
            {'paragraphs': 1000, 'styles': 30, 'fields': 25},
            {'paragraphs': 2000, 'styles': 40, 'fields': 35},
            {'paragraphs': 5000, 'styles': 50, 'fields': 50}
        ]
        
        for i, complexity in enumerate(complexities):
            test_name = f"validation_complexity_{i+1}"
            
            try:
                self.metrics.start_measurement(test_name)
                
                # Mock validation with complex document
                with patch('autoword.vnext.validator.DocumentValidator') as mock_validator_class:
                    mock_validator = Mock()
                    
                    # Create complex structure
                    structure = DocumentGenerator.generate_large_structure(
                        paragraph_count=complexity['paragraphs'],
                        style_count=complexity['styles'],
                        field_count=complexity['fields']
                    )
                    
                    # Mock validation result
                    mock_validation_result = Mock()
                    mock_validation_result.is_valid = True
                    mock_validation_result.errors = []
                    mock_validation_result.warnings = []
                    
                    mock_validator.validate_modifications.return_value = mock_validation_result
                    mock_validator.__enter__.return_value = mock_validator
                    mock_validator.__exit__.return_value = None
                    mock_validator_class.return_value = mock_validator
                    
                    # Simulate validation work
                    validator = mock_validator_class()
                    with validator:
                        result = validator.validate_modifications(structure, f"test_{i}.docx")
                        
                    # Simulate processing time based on complexity
                    time.sleep((complexity['paragraphs'] + complexity['styles'] + complexity['fields']) / 10000)
                    
                metrics = self.metrics.end_measurement(test_name)
                
                results['test_cases'].append({
                    'test_name': test_name,
                    'complexity': complexity,
                    'duration_seconds': metrics['duration_seconds'],
                    'memory_delta_mb': metrics['memory_delta_mb'],
                    'meets_threshold': metrics['duration_seconds'] <= self.thresholds['validation_max_seconds'],
                    'elements_per_second': sum(complexity.values()) / metrics['duration_seconds'] if metrics['duration_seconds'] > 0 else 0
                })
                
            except Exception as e:
                results['test_cases'].append({
                    'test_name': test_name,
                    'complexity': complexity,
                    'error': str(e),
                    'meets_threshold': False
                })
                
        return results   
     
    def benchmark_pipeline_performance(self) -> Dict[str, Any]:
        """Benchmark end-to-end pipeline performance."""
        results = {
            'test_cases': [],
            'summary': {},
            'threshold_compliance': {}
        }
        
        # Test different document sizes for full pipeline
        pipeline_tests = [
            {'size': 100, 'operations': 5, 'name': 'small_document'},
            {'size': 500, 'operations': 15, 'name': 'medium_document'},
            {'size': 1000, 'operations': 25, 'name': 'large_document'},
            {'size': 2000, 'operations': 40, 'name': 'very_large_document'}
        ]
        
        for test_config in pipeline_tests:
            test_name = f"pipeline_{test_config['name']}"
            
            try:
                self.metrics.start_measurement(test_name)
                
                # Mock full pipeline execution
                pipeline = VNextPipeline()
                
                with patch.multiple(pipeline,
                                  _setup_run_environment=Mock(),
                                  _extract_document=Mock(return_value=(
                                      DocumentGenerator.generate_large_structure(test_config['size']),
                                      Mock()
                                  )),
                                  _generate_plan=Mock(return_value=DocumentGenerator.generate_complex_plan(test_config['operations'])),
                                  _execute_plan=Mock(return_value=f"/modified_{test_config['name']}.docx"),
                                  _validate_modifications=Mock(return_value=Mock(is_valid=True)),
                                  _create_audit_trail=Mock(),
                                  _cleanup_run_environment=Mock()):
                    
                    pipeline.current_audit_dir = "/audit/test"
                    
                    # Simulate full pipeline processing
                    result = pipeline.process_document(f"test_{test_config['name']}.docx", "Test modification")
                    
                    # Simulate processing time based on document size and operations
                    processing_time = (test_config['size'] / 1000) + (test_config['operations'] / 100)
                    time.sleep(processing_time)
                    
                metrics = self.metrics.end_measurement(test_name)
                
                total_threshold = (self.thresholds['extraction_max_seconds'] + 
                                 self.thresholds['planning_max_seconds'] + 
                                 self.thresholds['execution_max_seconds'] + 
                                 self.thresholds['validation_max_seconds'])
                
                results['test_cases'].append({
                    'test_name': test_name,
                    'document_size': test_config['size'],
                    'operation_count': test_config['operations'],
                    'duration_seconds': metrics['duration_seconds'],
                    'memory_delta_mb': metrics['memory_delta_mb'],
                    'meets_threshold': metrics['duration_seconds'] <= total_threshold,
                    'throughput_paragraphs_per_second': test_config['size'] / metrics['duration_seconds'] if metrics['duration_seconds'] > 0 else 0
                })
                
            except Exception as e:
                results['test_cases'].append({
                    'test_name': test_name,
                    'document_size': test_config['size'],
                    'operation_count': test_config['operations'],
                    'error': str(e),
                    'meets_threshold': False
                })
                
        return results
        
    def benchmark_scalability(self) -> Dict[str, Any]:
        """Benchmark scalability with increasing document sizes."""
        results = {
            'test_cases': [],
            'scalability_analysis': {},
            'linear_scaling_assessment': {}
        }
        
        # Test scalability with exponentially increasing sizes
        sizes = [100, 200, 500, 1000, 2000, 5000]
        
        for size in sizes:
            test_name = f"scalability_{size}"
            
            try:
                self.metrics.start_measurement(test_name)
                
                # Mock processing with increasing document size
                structure = DocumentGenerator.generate_large_structure(
                    paragraph_count=size,
                    heading_count=size // 20,
                    style_count=min(50, size // 100),
                    field_count=min(20, size // 250)
                )
                
                # Simulate processing that scales with document size
                processing_time = size / 5000  # Linear scaling assumption
                time.sleep(processing_time)
                
                metrics = self.metrics.end_measurement(test_name)
                
                results['test_cases'].append({
                    'test_name': test_name,
                    'document_size': size,
                    'duration_seconds': metrics['duration_seconds'],
                    'memory_delta_mb': metrics['memory_delta_mb'],
                    'paragraphs_per_second': size / metrics['duration_seconds'] if metrics['duration_seconds'] > 0 else 0,
                    'memory_per_paragraph_kb': (metrics['memory_delta_mb'] * 1024) / size if size > 0 else 0
                })
                
            except Exception as e:
                results['test_cases'].append({
                    'test_name': test_name,
                    'document_size': size,
                    'error': str(e)
                })
                
        # Analyze scalability
        successful_tests = [t for t in results['test_cases'] if 'error' not in t]
        if len(successful_tests) >= 3:
            sizes = [t['document_size'] for t in successful_tests]
            durations = [t['duration_seconds'] for t in successful_tests]
            
            # Calculate scaling factor (should be close to 1.0 for linear scaling)
            scaling_factors = []
            for i in range(1, len(sizes)):
                size_ratio = sizes[i] / sizes[i-1]
                duration_ratio = durations[i] / durations[i-1]
                scaling_factor = duration_ratio / size_ratio
                scaling_factors.append(scaling_factor)
                
            results['scalability_analysis'] = {
                'average_scaling_factor': statistics.mean(scaling_factors),
                'scaling_consistency': statistics.stdev(scaling_factors) if len(scaling_factors) > 1 else 0,
                'linear_scaling': abs(statistics.mean(scaling_factors) - 1.0) < 0.5,  # Within 50% of linear
                'scaling_factors': scaling_factors
            }
            
        return results
        
    def benchmark_concurrency(self) -> Dict[str, Any]:
        """Benchmark concurrent processing performance."""
        results = {
            'test_cases': [],
            'concurrency_analysis': {},
            'thread_safety_assessment': {}
        }
        
        # Test different concurrency levels
        concurrency_levels = [1, 2, 4, 8]
        
        for concurrency in concurrency_levels:
            test_name = f"concurrency_{concurrency}_threads"
            
            try:
                self.metrics.start_measurement(test_name)
                
                # Mock concurrent processing
                def process_document_concurrent():
                    pipeline = VNextPipeline()
                    
                    with patch.multiple(pipeline,
                                      _setup_run_environment=Mock(),
                                      _extract_document=Mock(return_value=(Mock(), Mock())),
                                      _generate_plan=Mock(return_value=Mock()),
                                      _execute_plan=Mock(return_value="/modified.docx"),
                                      _validate_modifications=Mock(return_value=Mock(is_valid=True)),
                                      _create_audit_trail=Mock(),
                                      _cleanup_run_environment=Mock()):
                        
                        pipeline.current_audit_dir = "/audit/test"
                        result = pipeline.process_document("test.docx", "Test modification")
                        
                        # Simulate processing time
                        time.sleep(0.1)
                        return result
                        
                # Run concurrent processing
                with ThreadPoolExecutor(max_workers=concurrency) as executor:
                    futures = [executor.submit(process_document_concurrent) for _ in range(concurrency)]
                    results_list = [future.result() for future in as_completed(futures)]
                    
                metrics = self.metrics.end_measurement(test_name)
                
                results['test_cases'].append({
                    'test_name': test_name,
                    'concurrency_level': concurrency,
                    'duration_seconds': metrics['duration_seconds'],
                    'memory_delta_mb': metrics['memory_delta_mb'],
                    'successful_processes': len([r for r in results_list if r.status == "SUCCESS"]),
                    'throughput_processes_per_second': concurrency / metrics['duration_seconds'] if metrics['duration_seconds'] > 0 else 0
                })
                
            except Exception as e:
                results['test_cases'].append({
                    'test_name': test_name,
                    'concurrency_level': concurrency,
                    'error': str(e)
                })
                
        # Analyze concurrency efficiency
        successful_tests = [t for t in results['test_cases'] if 'error' not in t]
        if len(successful_tests) >= 2:
            single_thread = next((t for t in successful_tests if t['concurrency_level'] == 1), None)
            if single_thread:
                results['concurrency_analysis'] = {
                    'baseline_duration': single_thread['duration_seconds'],
                    'efficiency_by_threads': {}
                }
                
                for test in successful_tests:
                    if test['concurrency_level'] > 1:
                        expected_duration = single_thread['duration_seconds'] / test['concurrency_level']
                        actual_duration = test['duration_seconds']
                        efficiency = expected_duration / actual_duration if actual_duration > 0 else 0
                        
                        results['concurrency_analysis']['efficiency_by_threads'][test['concurrency_level']] = {
                            'expected_duration': expected_duration,
                            'actual_duration': actual_duration,
                            'efficiency_ratio': efficiency
                        }
                        
        return results
        
    def benchmark_memory_usage(self) -> Dict[str, Any]:
        """Benchmark memory usage patterns."""
        results = {
            'test_cases': [],
            'memory_analysis': {},
            'memory_leak_assessment': {}
        }
        
        # Test memory usage with different scenarios
        memory_tests = [
            {'name': 'small_document', 'size': 100, 'iterations': 10},
            {'name': 'medium_document', 'size': 500, 'iterations': 5},
            {'name': 'large_document', 'size': 1000, 'iterations': 3},
            {'name': 'repeated_processing', 'size': 200, 'iterations': 20}
        ]
        
        for test_config in memory_tests:
            test_name = f"memory_{test_config['name']}"
            
            try:
                initial_memory = self._get_current_memory()
                memory_snapshots = [initial_memory]
                
                self.metrics.start_measurement(test_name)
                
                # Run multiple iterations to test for memory leaks
                for iteration in range(test_config['iterations']):
                    # Mock document processing
                    structure = DocumentGenerator.generate_large_structure(test_config['size'])
                    
                    # Simulate processing
                    time.sleep(0.01)  # Small delay to simulate work
                    
                    # Take memory snapshot
                    current_memory = self._get_current_memory()
                    memory_snapshots.append(current_memory)
                    
                    # Force garbage collection
                    import gc
                    gc.collect()
                    
                metrics = self.metrics.end_measurement(test_name)
                
                # Analyze memory usage
                memory_growth = memory_snapshots[-1] - memory_snapshots[0]
                peak_memory = max(memory_snapshots)
                average_memory = statistics.mean(memory_snapshots)
                
                # Check for memory leaks (consistent growth)
                if len(memory_snapshots) > 5:
                    recent_growth = memory_snapshots[-1] - memory_snapshots[-5]
                    memory_leak_suspected = recent_growth > 50  # More than 50MB growth in last 5 iterations
                else:
                    memory_leak_suspected = False
                    
                results['test_cases'].append({
                    'test_name': test_name,
                    'iterations': test_config['iterations'],
                    'document_size': test_config['size'],
                    'duration_seconds': metrics['duration_seconds'],
                    'initial_memory_mb': memory_snapshots[0],
                    'final_memory_mb': memory_snapshots[-1],
                    'peak_memory_mb': peak_memory,
                    'average_memory_mb': average_memory,
                    'memory_growth_mb': memory_growth,
                    'memory_leak_suspected': memory_leak_suspected,
                    'meets_memory_threshold': peak_memory <= self.thresholds['memory_max_mb'],
                    'memory_snapshots': memory_snapshots
                })
                
            except Exception as e:
                results['test_cases'].append({
                    'test_name': test_name,
                    'iterations': test_config['iterations'],
                    'document_size': test_config['size'],
                    'error': str(e)
                })
                
        # Overall memory analysis
        successful_tests = [t for t in results['test_cases'] if 'error' not in t]
        if successful_tests:
            total_memory_growth = sum(t['memory_growth_mb'] for t in successful_tests)
            suspected_leaks = [t for t in successful_tests if t['memory_leak_suspected']]
            
            results['memory_analysis'] = {
                'total_memory_growth_mb': total_memory_growth,
                'average_memory_growth_mb': total_memory_growth / len(successful_tests),
                'suspected_memory_leaks': len(suspected_leaks),
                'memory_efficiency_score': max(0, 1.0 - (total_memory_growth / 1000))  # Penalize excessive growth
            }
            
        return results
        
    def generate_performance_assessment(self) -> Dict[str, Any]:
        """Generate overall performance assessment."""
        assessment = {
            'overall_performance_score': 0.0,
            'performance_categories': {},
            'threshold_violations': [],
            'performance_recommendations': [],
            'scalability_rating': 'Unknown',
            'memory_efficiency_rating': 'Unknown'
        }
        
        # This would be populated after running all benchmarks
        # For now, return a template structure
        
        return assessment
        
    def save_benchmark_results(self, results: Dict[str, Any]):
        """Save benchmark results to files."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Save JSON results
        json_path = self.output_dir / f"performance_results_{timestamp}.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, default=str)
            
        # Save performance report
        report_path = self.output_dir / f"performance_report_{timestamp}.html"
        html_report = self._generate_html_report(results)
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(html_report)
            
        # Save CSV summary
        csv_path = self.output_dir / f"performance_summary_{timestamp}.csv"
        self._generate_csv_summary(results, csv_path)
        
    def _generate_html_report(self, results: Dict[str, Any]) -> str:
        """Generate HTML performance report."""
        html = f"""
<!DOCTYPE html>
<html>
<head>
    <title>AutoWord vNext Performance Benchmark Report</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        .summary {{ background-color: #f0f0f0; padding: 15px; border-radius: 5px; margin: 20px 0; }}
        .benchmark-section {{ margin: 30px 0; }}
        .test-case {{ margin: 10px 0; padding: 10px; border: 1px solid #ddd; }}
        .pass {{ background-color: #d4edda; }}
        .fail {{ background-color: #f8d7da; }}
        table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; }}
        .chart {{ margin: 20px 0; }}
    </style>
</head>
<body>
    <h1>AutoWord vNext Performance Benchmark Report</h1>
    <p>Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
    
    <div class="summary">
        <h2>Performance Summary</h2>
        <p>Total Benchmark Categories: {len([k for k in results.keys() if k != 'overall_assessment'])}</p>
        <p>Performance Thresholds: {self.thresholds}</p>
    </div>
    
    <div class="benchmark-section">
        <h2>Extraction Performance</h2>
        {self._format_benchmark_section_html(results.get('extraction_benchmarks', {}))}
    </div>
    
    <div class="benchmark-section">
        <h2>Planning Performance</h2>
        {self._format_benchmark_section_html(results.get('planning_benchmarks', {}))}
    </div>
    
    <div class="benchmark-section">
        <h2>Execution Performance</h2>
        {self._format_benchmark_section_html(results.get('execution_benchmarks', {}))}
    </div>
    
    <div class="benchmark-section">
        <h2>Validation Performance</h2>
        {self._format_benchmark_section_html(results.get('validation_benchmarks', {}))}
    </div>
    
    <div class="benchmark-section">
        <h2>Pipeline Performance</h2>
        {self._format_benchmark_section_html(results.get('pipeline_benchmarks', {}))}
    </div>
    
    <div class="benchmark-section">
        <h2>Scalability Analysis</h2>
        {self._format_scalability_section_html(results.get('scalability_benchmarks', {}))}
    </div>
    
    <div class="benchmark-section">
        <h2>Memory Usage Analysis</h2>
        {self._format_memory_section_html(results.get('memory_benchmarks', {}))}
    </div>
    
</body>
</html>
        """
        return html
        
    def _format_benchmark_section_html(self, benchmark_data: Dict[str, Any]) -> str:
        """Format benchmark section for HTML report."""
        if not benchmark_data or 'test_cases' not in benchmark_data:
            return "<p>No benchmark data available</p>"
            
        html = "<table><tr><th>Test</th><th>Duration (s)</th><th>Memory (MB)</th><th>Threshold Met</th></tr>"
        
        for test_case in benchmark_data['test_cases']:
            threshold_class = "pass" if test_case.get('meets_threshold', False) else "fail"
            html += f"""
            <tr class="{threshold_class}">
                <td>{test_case.get('test_name', 'Unknown')}</td>
                <td>{test_case.get('duration_seconds', 0):.3f}</td>
                <td>{test_case.get('memory_delta_mb', 0):.2f}</td>
                <td>{'✓' if test_case.get('meets_threshold', False) else '✗'}</td>
            </tr>
            """
            
        html += "</table>"
        
        if 'summary' in benchmark_data:
            summary = benchmark_data['summary']
            html += f"""
            <div class="summary">
                <h4>Summary</h4>
                <p>Average Duration: {summary.get('average_duration', 0):.3f}s</p>
                <p>Threshold Violations: {summary.get('threshold_violations', 0)}</p>
            </div>
            """
            
        return html
        
    def _format_scalability_section_html(self, scalability_data: Dict[str, Any]) -> str:
        """Format scalability section for HTML report."""
        if not scalability_data:
            return "<p>No scalability data available</p>"
            
        html = ""
        
        if 'scalability_analysis' in scalability_data:
            analysis = scalability_data['scalability_analysis']
            html += f"""
            <div class="summary">
                <h4>Scalability Analysis</h4>
                <p>Average Scaling Factor: {analysis.get('average_scaling_factor', 0):.2f}</p>
                <p>Linear Scaling: {'✓' if analysis.get('linear_scaling', False) else '✗'}</p>
                <p>Scaling Consistency: {analysis.get('scaling_consistency', 0):.3f}</p>
            </div>
            """
            
        return html
        
    def _format_memory_section_html(self, memory_data: Dict[str, Any]) -> str:
        """Format memory section for HTML report."""
        if not memory_data:
            return "<p>No memory data available</p>"
            
        html = ""
        
        if 'memory_analysis' in memory_data:
            analysis = memory_data['memory_analysis']
            html += f"""
            <div class="summary">
                <h4>Memory Analysis</h4>
                <p>Total Memory Growth: {analysis.get('total_memory_growth_mb', 0):.2f} MB</p>
                <p>Suspected Memory Leaks: {analysis.get('suspected_memory_leaks', 0)}</p>
                <p>Memory Efficiency Score: {analysis.get('memory_efficiency_score', 0):.2f}</p>
            </div>
            """
            
        return html
        
    def _generate_csv_summary(self, results: Dict[str, Any], csv_path: Path):
        """Generate CSV summary of performance results."""
        import csv
        
        with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Category', 'Test Name', 'Duration (s)', 'Memory Delta (MB)', 'Threshold Met', 'Notes'])
            
            for category, category_data in results.items():
                if category == 'overall_assessment' or 'test_cases' not in category_data:
                    continue
                    
                for test_case in category_data['test_cases']:
                    writer.writerow([
                        category,
                        test_case.get('test_name', ''),
                        test_case.get('duration_seconds', 0),
                        test_case.get('memory_delta_mb', 0),
                        test_case.get('meets_threshold', False),
                        test_case.get('error', '')
                    ])
                    
    def _get_current_memory(self) -> float:
        """Get current memory usage in MB."""
        try:
            process = psutil.Process()
            return process.memory_info().rss / 1024 / 1024
        except Exception:
            return 0.0


# Main execution and test runner
class PerformanceTestRunner:
    """Main performance test runner."""
    
    def __init__(self):
        self.benchmark_suite = PerformanceBenchmarkSuite()
        
    def run_performance_tests(self) -> Dict[str, Any]:
        """Run all performance tests."""
        print("Starting AutoWord vNext Performance Benchmarks...")
        
        results = self.benchmark_suite.run_all_benchmarks()
        
        # Print summary
        self._print_performance_summary(results)
        
        return results
        
    def _print_performance_summary(self, results: Dict[str, Any]):
        """Print performance summary to console."""
        print("\n=== AutoWord vNext Performance Benchmark Results ===")
        
        for category, category_data in results.items():
            if category == 'overall_assessment':
                continue
                
            print(f"\n{category.replace('_', ' ').title()}:")
            
            if 'test_cases' in category_data:
                successful_tests = [t for t in category_data['test_cases'] if 'error' not in t]
                failed_tests = [t for t in category_data['test_cases'] if 'error' in t]
                
                print(f"  Successful Tests: {len(successful_tests)}")
                print(f"  Failed Tests: {len(failed_tests)}")
                
                if successful_tests:
                    durations = [t['duration_seconds'] for t in successful_tests]
                    print(f"  Average Duration: {statistics.mean(durations):.3f}s")
                    print(f"  Max Duration: {max(durations):.3f}s")
                    
                    threshold_violations = [t for t in successful_tests if not t.get('meets_threshold', True)]
                    print(f"  Threshold Violations: {len(threshold_violations)}")
                    
        print(f"\nDetailed results saved to: {self.benchmark_suite.output_dir}")


if __name__ == "__main__":
    runner = PerformanceTestRunner()
    results = runner.run_performance_tests()