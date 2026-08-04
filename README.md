OPENMOON/
│
├── assets/                     # 로고, 직인, 아이콘 등 (보안상 올리지 않음)
│
├── quotation_drafts/           # 생성된 견적서
│
├── quotation_files/            # 기존 견적서 샘플(xlsx) (보안상 올리지 않음)
│
├── .env                        # 환경설정 (GitHub 업로드 X) (보안상 올리지 않음)
│
├── 단가표.xlsx                 # 회사 단가표 원본 (보안상 올리지 않음)
│
├── build_price_db.py           # 단가표 → price_table.db 생성
├── build_quote_db.py           # 견적서 샘플 → quotation_history.db 생성
├── build_price_db_check.py     # 단가 DB 검증
│
├── ex.xlsx                     # 견적서 템플릿
│
├── price_engine.py             # 가격 검색 엔진
├── price_table.db              # 단가 DB (보안상 올리지 않음)
├── quotation_history.db        # 견적서 DB (보안상 올리지 않음)
│
├── read_quote_db.py            # 견적 DB 확인용
│
├── requirements.txt
│
├── YullinMoon_Ver3.py          # 메인 실행 파일
│
└── README.md
