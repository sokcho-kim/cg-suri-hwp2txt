"""
PDF 서비스 레이어  
PDF 관련 비즈니스 워크플로우를 조합하여 처리
"""

import fitz
import base64
import io
from typing import Dict, Any
from app.core.pdf import PDFMasker, PDFConverter, TextModifier
from app.core.pdf.direct_edit_filter import DirectEditFilter
from app.core.llm import LLMClient
from config import logging_config


logger = logging_config.get_logger(__name__)

class PDFService:
    """PDF 관련 애플리케이션 서비스"""
    
    def __init__(self):
        self.text_modifier = TextModifier()
        self.pdf_masker = PDFMasker()
        self.pdf_converter = PDFConverter()
        self.llm_client = LLMClient()
        self.direct_edit_filter = DirectEditFilter()
    
    def process_pdf_modification(self, pdf_data: bytes, modifications: Dict) -> bytes:
        """
        PDF 텍스트 수정 워크플로우 (메모리 기반) - 분리된 수정사항 처리
        
        Args:
            pdf_data: PDF 바이트 데이터
            modifications: 분리된 수정사항 딕셔너리
                         {"directEdits": [...], "toggleMasks": [...]}
                         또는 기존 호환을 위한 리스트
            
        Returns:
            bytes: 수정된 PDF 바이트 데이터
        """
        logger.info(f"PDF 수정 시작")
        
        direct_edits = modifications.get('directEdits', [])
        toggle_masks = modifications.get('toggleMasks', [])
        
        logger.info(f"분리된 수정사항 처리: 직접편집 {len(direct_edits)}개, 토글마스킹 {len(toggle_masks)}개")
        
        # 두 가지 수정사항을 순차적으로 처리
        modified_bytes = self._process_separated_modifications(pdf_data, direct_edits, toggle_masks)
        
        logger.info(f"PDF 수정 완료: 결과 크기: {len(modified_bytes)} bytes")
        return modified_bytes
    
    def _process_legacy_modifications(self, pdf_data: bytes, modifications: list) -> bytes:
        """
        기존 방식 호환을 위한 수정사항 처리
        
        Args:
            pdf_data: PDF 바이트 데이터
            modifications: 기존 방식 수정사항 리스트
            
        Returns:
            bytes: 수정된 PDF 바이트 데이터
        """
        logger.debug("기존 방식으로 수정사항 처리")
        return self.text_modifier.modify_pdf(pdf_data, modifications)
    
    def _process_separated_modifications(self, pdf_data: bytes, direct_edits: list, toggle_masks: list) -> bytes:
        """
        분리된 수정사항들을 순차적으로 처리
        
        Args:
            pdf_data: PDF 바이트 데이터
            direct_edits: 직접 수정사항 리스트
            toggle_masks: 토글 마스킹 수정사항 리스트
            
        Returns:
            bytes: 수정된 PDF 바이트 데이터
        """
        current_pdf_data = pdf_data
        
        # Step 1: 직접 수정사항 처리 (우선순위 높음)
        if direct_edits:
            logger.info(f"직접 수정사항 처리 시작: {len(direct_edits)}개")
            current_pdf_data = self._process_direct_edits(current_pdf_data, direct_edits)
            logger.info("직접 수정사항 처리 완료")

        # Step 2: 토글 마스킹 처리 (텍스트 검증 포함)
        if toggle_masks:
            logger.info(f"토글 마스킹 처리 시작: {len(toggle_masks)}개")
            # DirectEditFilter로 토글도 검증
            validated_toggles = self.direct_edit_filter.filter_toggle_modifications(current_pdf_data, toggle_masks)
            if validated_toggles:
                current_pdf_data = self.text_modifier.modify_pdf(current_pdf_data, validated_toggles)
        logger.info("토글 마스킹 처리 완료")

        return current_pdf_data
    
    def _process_direct_edits(self, pdf_data: bytes, direct_edits: list) -> bytes:
        """
        직접 수정사항 처리 (char 단위 직접 처리)
        
        Args:
            pdf_data: PDF 바이트 데이터
            direct_edits: 직접 수정사항 리스트 (char 단위)
            
        Returns:
            bytes: 수정된 PDF 바이트 데이터
        """
        logger.debug(f"직접 수정사항 처리: {len(direct_edits)}개 char 단위")
        
        if not direct_edits:
            logger.debug("직접 수정사항이 없음, 원본 반환")
            return pdf_data
        
        # char 단위로 바로 TextModifier에 전달
        return self.text_modifier.modify_pdf(pdf_data, direct_edits)
    
    def _process_toggle_masks(self, pdf_data: bytes, toggle_masks: list) -> bytes:
        """
        토글 마스킹 수정사항 처리 (기존 방식)
        
        Args:
            pdf_data: PDF 바이트 데이터
            toggle_masks: 토글 마스킹 수정사항 리스트 (char 단위)
            
        Returns:
            bytes: 수정된 PDF 바이트 데이터
        """
        logger.debug(f"토글 마스킹 처리: {len(toggle_masks)}개 char 수정사항")
        
        # 토글 마스킹은 이미 char 단위로 정확한 좌표를 가지고 있으므로
        # 바로 기존 modifier로 처리
        return self.text_modifier.modify_pdf(pdf_data, toggle_masks)
    
    def process_pdf_masking(self, file_data: bytes, filename: str = None, masking_settings: Dict = None) -> Dict[str, Any]:
        """
        PDF 마스킹 처리 워크플로우 (순수 메모리 기반)
        
        Args:
            file_data: 업로드된 파일의 바이트 데이터
            filename: 업로드된 파일명 (확장자 확인용)
            masking_settings: 마스킹 설정 (어떤 항목을 마스킹할지, 기호 사용할지 등)
            
        Returns:
            Dict: 마스킹 결과 (Base64 PDF 데이터, 패턴)
        """
        try:
            masking_settings = masking_settings or {}
            logger.info(f"마스킹 시작: 파일명={filename}, 크기={len(file_data)} bytes, 설정={masking_settings}")
            
            # 1. 파일 형식 검사 및 PDF로 변환
            pdf_bytes = self.pdf_converter.convert_to_pdf(file_data, filename)
            logger.debug(f"PDF 변환 완료: {len(pdf_bytes)} bytes")
            
            # 2. PDF 바이트에서 텍스트 추출
            extracted_text = self._extract_text_from_pdf(pdf_bytes)
            logger.debug(f"텍스트 추출 완료: {len(extracted_text)} 문자")

            # 3. 활성화된 항목들만 추출
            enabled_items = {k: v for k, v in masking_settings.items() if v.get('enabled', False)}
            
            # 4. LLM API 요청으로 개인정보 패턴 추출 (활성화된 항목만)
            patterns = self.llm_client.request_llm_masking(extracted_text, enabled_items)
            logger.debug(f"LLM 패턴 추출 완료: {patterns}")
            
            # 빈 패턴들 제거
            filtered_patterns = {k: [p for p in v if p.strip()] for k, v in patterns.items() if v}
            
            # 마스킹할 개인정보가 없는 경우 원본 반환
            if not any(filtered_patterns.values()):
                logger.warning("문서에서 마스킹할 개인정보를 찾을 수 없음")
                # 원본 PDF를 Base64로 변환하여 반환
                original_pdf_base64 = base64.b64encode(pdf_bytes).decode('utf-8')
                return {
                    'pdf_data': original_pdf_base64,
                    'patterns': {}
                }
            
            # 5. PDF 바이트에서 마스킹 맵 생성 (심볼이 있는 항목만, 실제 마스킹 처리는 하지 않음)
            original_pdf_bytes, masking_map = self.pdf_masker.mask_pdf_from_bytes(
                pdf_bytes, filtered_patterns, masking_settings=enabled_items
            )
            
            # 6. 마스킹 맵에서 실제 발견된 패턴들만 추출하여 최종 patterns 생성
            final_patterns = self._extract_found_patterns_from_masking_map(filtered_patterns, masking_map)
            
            # 7. 원본 PDF 바이트를 Base64로 변환
            pdf_base64 = base64.b64encode(original_pdf_bytes).decode('utf-8')
            
            found_pattern_count = sum(len(v) for v in final_patterns.values())
            logger.info(f"마스킹 처리 완료: 실제 발견된 패턴 {found_pattern_count}개")
            
            return {
                'pdf_data': pdf_base64,  # Base64 인코딩된 원본 PDF 데이터
                'patterns': final_patterns,  # 실제 발견된 패턴들만
                'masking_map': masking_map  # 마스킹 기호 좌표 정보
            }
                
        except Exception as e:
            logger.error(f"마스킹 서비스 처리 에러: {e}")
            raise Exception(f"마스킹 서비스 처리 에러: {e}")
    
    def _extract_found_patterns_from_masking_map(self, original_patterns: Dict, masking_map: Dict) -> Dict:
        """
        마스킹 맵에서 실제 발견된 패턴들만 추출
        
        Args:
            original_patterns: LLM에서 추출한 원본 패턴들
            masking_map: 실제 PDF에서 발견된 패턴들의 좌표 정보
            
        Returns:
            Dict: 실제 발견된 패턴들만 포함한 딕셔너리
        """
        final_patterns = {}
        
        for pattern_type, pattern_list in original_patterns.items():
            found_patterns = []
            
            for pattern in pattern_list:
                # 마스킹 맵에 해당 패턴이 있는지 확인
                if pattern in masking_map and masking_map[pattern]:
                    found_patterns.append(pattern)
                    logger.debug(f"패턴 '{pattern}' - PDF에서 발견됨")
                else:
                    logger.debug(f"패턴 '{pattern}' - PDF에서 발견되지 않음")
            
            # 발견된 패턴이 있는 경우만 최종 결과에 포함
            if found_patterns:
                final_patterns[pattern_type] = found_patterns
        
        logger.info(f"패턴 필터링 결과: {len(original_patterns)} -> {len(final_patterns)} 항목")
        return final_patterns
    
    def _extract_text_from_pdf(self, pdf_bytes: bytes) -> str:
        """PDF 바이트에서 텍스트 추출"""
        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            extracted_text = ""
            
            for page in doc:
                extracted_text += page.get_text()
            
            doc.close()
            return extracted_text
            
        except Exception as e:
            raise ValueError(f"텍스트 추출 실패: {str(e)}")