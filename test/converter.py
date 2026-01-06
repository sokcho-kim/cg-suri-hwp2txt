"""
PDF 변환기 모듈
DOCX, HWP 파일을 PDF로 변환하는 기능을 제공
"""

import os
import uuid
import time
import tempfile
import win32com.client as win32
import pythoncom
# from config import logging_config


class PDFConverter:
    """DOCX 또는 HWP 파일을 PDF로 변환하는 클래스"""

    def __init__(self):
        """생성자"""
        self.instance_id = str(int(time.time())) + "_" + str(uuid.uuid4())[:8]

    def convert_to_pdf(self, file_data: bytes, filename: str = None) -> bytes:
        """
        파일을 PDF로 변환 (파일 형식 검증 포함)
        
        Args:
            file_data: 파일의 바이트 데이터
            filename: 파일명 (확장자 확인용)
            
        Returns:
            bytes: PDF 바이트 데이터
            
        Raises:
            ValueError: 지원하지 않는 파일 형식
        """
        # 파일 확장자 결정 (파일명 우선, 바이트 헤더 보조)
        file_ext = self._determine_file_extension(file_data, filename)
        
        if file_ext == '.unknown':
            raise ValueError("지원하지 않는 파일 형식입니다.")
        
        # PDF 변환 처리
        if file_ext == '.pdf':
            return file_data
        # 리눅스 서버(테스트 배포)에서 주석처리
        # elif file_ext in ['.doc', '.docx', '.hwp', '.hwpx']:
        #     return self._convert_office_to_pdf_bytes(file_data, file_ext)
        else:
            raise ValueError(f"지원하지 않는 파일 형식: {file_ext}")

    def _determine_file_extension(self, file_data: bytes, filename: str = None) -> str:
        """
        파일 확장자 결정 (파일명 우선, 바이트 헤더 보조)
        
        Args:
            file_data: 파일의 바이트 데이터
            filename: 파일명
            
        Returns:
            str: 결정된 파일 확장자
        """
        file_ext = None
        
        if filename:
            # 파일명에서 확장자 추출 (신뢰도 높음)
            file_ext = os.path.splitext(filename.lower())[1]
        
        # 파일명 확장자가 없거나 지원하지 않는 형식인 경우 바이트 헤더로 추정
        if not file_ext or file_ext not in ['.pdf', '.doc', '.docx', '.hwp', '.hwpx']:
            file_ext = self.get_file_extension_from_bytes(file_data)
        
        return file_ext

    def _convert_office_to_pdf_bytes(self, file_bytes: bytes, file_extension: str) -> bytes:
        """
        오피스 파일 바이트 데이터를 PDF 바이트로 변환
        
        Args:
            file_bytes: 원본 파일의 바이트 데이터
            file_extension: 파일 확장자 (.docx, .hwp 등)
            
        Returns:
            bytes: 변환된 PDF 바이트 데이터
            
        Raises:
            Exception: 파일 변환 중 오류 발생 시
        """
        file_ext = file_extension.lower()
        
        # 임시 파일 생성 (COM 객체 요구사항)
        temp_input_path = None
        temp_output_path = None
        
        try:
            # 입력 파일 임시 생성
            with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as temp_input:
                temp_input.write(file_bytes)
                temp_input_path = temp_input.name
            
            # 출력 파일 경로 생성
            temp_output_path = tempfile.mktemp(suffix=".pdf")
            
            # 파일 형식에 따른 변환
            if file_ext in ['.doc', '.docx']:
                self._convert_word_to_pdf(temp_input_path, temp_output_path)
            elif file_ext in ['.hwp', '.hwpx']:
                self._convert_hwp_to_pdf(temp_input_path, temp_output_path)
            else:
                raise ValueError(f"지원하지 않는 파일 형식: {file_ext}")
            
            # 변환된 PDF를 바이트로 읽기
            with open(temp_output_path, 'rb') as f:
                pdf_bytes = f.read()

            print(f"파일 변환 완료: {file_ext} → PDF ({len(pdf_bytes)} bytes)")
            return pdf_bytes
            
        except Exception as e:
            raise Exception(f"파일 변환 실패: {str(e)}")
        finally:
            # 임시 파일 정리
            self._cleanup_temp_files([temp_input_path, temp_output_path])

    def _convert_word_to_pdf(self, word_path: str, pdf_path: str):
        """Word 문서를 PDF로 변환"""
        pythoncom.CoInitialize()
        try:
            word = win32.gencache.EnsureDispatch("Word.Application")
            word.Visible = False

            doc = word.Documents.Open(word_path)
            doc.SaveAs(pdf_path, FileFormat=17)  # 17은 PDF 포맷
            
            doc.Close()
            word.Quit()
            
            # 파일 생성 확인
            if not os.path.exists(pdf_path) or os.path.getsize(pdf_path) == 0:
                raise Exception("Word PDF 변환 실패: 출력 파일이 생성되지 않음")
            
        except Exception as e:
            raise Exception(f"Word PDF 변환 실패: {e}")
        finally:
            pythoncom.CoUninitialize()

    def _convert_hwp_to_pdf(self, hwp_path: str, pdf_path: str):
        """HWP 문서를 PDF로 변환"""
        pythoncom.CoInitialize()
        
        try:
            hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
            hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')

            hwp.Open(hwp_path)
            
            # PrintToPDF 액션 생성 및 설정
            action = hwp.CreateAction("PrintToPDF")
            pSet = action.CreateSet()
            action.GetDefault(pSet)
            
            # PDF 변환 옵션 설정
            pSet.SetItem("PrintMethod", 0)
            pSet.SetItem("PrintPageOption", 1)
            pSet.SetItem("FileName", pdf_path)
            
            # PDF로 변환 실행
            action.Execute(pSet)

            # 파일 생성 완료 대기 (최대 30초)
            import time
            max_wait_time = 30
            wait_time = 0
            
            while wait_time < max_wait_time:
                if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                    try:
                        # 파일이 완전히 쓰여졌는지 확인
                        with open(pdf_path, 'rb') as f:
                            f.read(1)
                        break
                    except:
                        pass
                
                time.sleep(0.5)
                wait_time += 0.5
            
            hwp.Quit()
            
            # 변환 성공 확인
            if not os.path.exists(pdf_path) or os.path.getsize(pdf_path) == 0:
                raise Exception("HWP PDF 변환 실패: 출력 파일이 생성되지 않음")
        
        except Exception as e:
            raise Exception(f"HWP PDF 변환 실패: {e}")
        finally:    
            pythoncom.CoUninitialize()

    def _cleanup_temp_files(self, file_paths: list):
        """임시 파일들 정리"""
        for file_path in file_paths:
            if file_path and os.path.exists(file_path):
                try:
                    os.unlink(file_path)
                    print(f"임시 파일 삭제: {file_path}")
                except Exception as e:
                    print(f"임시 파일 삭제 실패: {file_path} - {e}")

    def is_pdf_file(self, file_bytes: bytes) -> bool:
        """
        바이트 데이터가 PDF 파일인지 확인
        
        Args:
            file_bytes: 확인할 파일의 바이트 데이터
            
        Returns:
            bool: PDF 파일 여부
        """
        # PDF 파일은 %PDF- 로 시작
        return file_bytes.startswith(b'%PDF-')

    def get_file_extension_from_bytes(self, file_bytes: bytes) -> str:
        """
        바이트 데이터에서 파일 확장자 추정
        
        Args:
            file_bytes: 파일의 바이트 데이터
            
        Returns:
            str: 추정된 파일 확장자
        """
        # PDF 파일 확인
        if file_bytes.startswith(b'%PDF-'):
            return '.pdf'
        
        # DOCX 파일 확인 (ZIP 압축 형태)
        if file_bytes.startswith(b'PK\x03\x04'):
            # DOCX는 ZIP 파일이므로 더 자세한 확인 필요
            if b'word/' in file_bytes[:1000] or b'[Content_Types].xml' in file_bytes[:1000]:
                return '.docx'
        
        # DOC 파일 확인
        if file_bytes.startswith(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'):
            return '.doc'
        
        # HWP 파일 확인
        if file_bytes.startswith(b'HWP Document File'):
            return '.hwp'
        
        # 기본값
        return '.unknown'
