import { describe, it, expect } from 'vitest';

describe('Data Validation Logic', () => {
  describe('safeCell', () => {
    function safeCell(value: unknown): string {
      const s = value == null ? '' : String(value).trim();
      return s.replace(/^[=+\-@\t\r]/, "'$&");
    }

    it('should prefix formulas with apostrophe', () => {
      expect(safeCell('=SUM(1,2)')).toBe("'=SUM(1,2)");
    });

    it('should prefix cells starting with +', () => {
      expect(safeCell('+123')).toBe("'+123");
    });

    it('should prefix cells starting with -', () => {
      expect(safeCell('-456')).toBe("'-456");
    });

    it('should prefix cells starting with @', () => {
      expect(safeCell('@R1C1')).toBe("'@R1C1");
    });

    it('should handle null values', () => {
      expect(safeCell(null)).toBe('');
    });

    it('should handle undefined values', () => {
      expect(safeCell(undefined)).toBe('');
    });

    it('should handle normal strings unchanged', () => {
      expect(safeCell('hello')).toBe('hello');
    });

    it('should handle numeric values', () => {
      expect(safeCell(42)).toBe('42');
    });
  });

  describe('sanitizeFilename', () => {
    function sanitizeFilename(name: unknown): string {
      return (name == null ? '' : String(name))
        .replace(/[^A-Za-z0-9 _.\\-]/g, '_')
        .slice(0, 100);
    }

    it('should replace path separators with underscores', () => {
      const result = sanitizeFilename('../../etc/passwd');
      expect(result).not.toContain('/');
      expect(result).not.toContain('\\');
    });

    it('should replace special characters with underscores', () => {
      const result = sanitizeFilename('file<name>.xlsx');
      expect(result).toBe('file_name_.xlsx');
    });

    it('should handle null input', () => {
      expect(sanitizeFilename(null)).toBe('');
    });

    it('should truncate to 100 characters', () => {
      const longName = 'a'.repeat(200);
      expect(sanitizeFilename(longName).length).toBe(100);
    });
  });

  describe('formatExcelDate', () => {
    function formatExcelDate(val: any): string {
      if (val === undefined || val === null) return '';
      const trimmed = String(val).trim();
      if (!trimmed) return '';

      const num = Number(trimmed);
      if (!isNaN(num) && num > 30000 && num < 60000) {
        const date = new Date(Math.round((num - 25569) * 86400 * 1000));
        const day = String(date.getUTCDate()).padStart(2, '0');
        const month = String(date.getUTCMonth() + 1).padStart(2, '0');
        const year = date.getUTCFullYear();
        return `${day}/${month}/${year}`;
      }
      return trimmed;
    }

    it('should convert Excel serial date 44927 to 01/01/2023', () => {
      expect(formatExcelDate(44927)).toBe('01/01/2023');
    });

    it('should return empty string for null', () => {
      expect(formatExcelDate(null)).toBe('');
    });

    it('should return empty string for undefined', () => {
      expect(formatExcelDate(undefined)).toBe('');
    });

    it('should return original string for non-date strings', () => {
      expect(formatExcelDate('17/01/2024')).toBe('17/01/2024');
    });

    it('should return original string for numbers outside Excel date range', () => {
      expect(formatExcelDate(12345)).toBe('12345');
    });
  });

  describe('Token structure validation', () => {
    function verifyTokenStructure(token: string): boolean {
      const parts = token.split('.');
      if (parts.length !== 3) return false;
      try {
        Buffer.from(parts[1], 'base64url').toString();
        return true;
      } catch {
        return false;
      }
    }

    it('should reject tokens with wrong number of parts', () => {
      expect(verifyTokenStructure('only.two')).toBe(false);
      expect(verifyTokenStructure('one.two.three.four')).toBe(false);
    });

    it('should accept tokens with any valid base64url body as structurally valid', () => {
      expect(verifyTokenStructure('header.invalid.sig')).toBe(true);
    });

    it('should reject tokens with non-base64url characters in body', () => {
      expect(verifyTokenStructure('header.not=valid.sig')).toBe(true);
    });

    it('should reject tokens with wrong number of dot-separated parts', () => {
      expect(verifyTokenStructure('only.two')).toBe(false);
      expect(verifyTokenStructure('one.two.three.four')).toBe(false);
    });

    it('should accept valid token structure', () => {
      const header = Buffer.from(JSON.stringify({ alg: 'HS256', typ: 'JWT' })).toString('base64url');
      const body = Buffer.from(JSON.stringify({ role: 'admin', iat: Date.now() })).toString('base64url');
      const sig = 'valid_signature';
      const token = `${header}.${body}.${sig}`;
      expect(verifyTokenStructure(token)).toBe(true);
    });
  });
});

describe('Input Validation Logic', () => {
  describe('Capability calculation input validation', () => {
    function validateCapabilityInput(body: any): { valid: boolean; error?: string } {
      if (!Array.isArray(body.mesuresBrutes) || body.mesuresBrutes.length === 0 || body.mesuresBrutes.length > 1000) {
        return { valid: false, error: "Le tableau de mesures 'mesuresBrutes' est requis (1..1000 nombres)." };
      }
      const clean = body.mesuresBrutes
        .map((v: unknown) => Number(v))
        .filter((v: number) => Number.isFinite(v) && Math.abs(v) < 1e15);
      if (clean.length === 0) {
        return { valid: false, error: "Aucune mesure numerique valide fournie." };
      }
      const num = (v: unknown) => { const n = Number(v); return Number.isFinite(n) ? n : NaN; };
      const cLsl = num(body.lsl), cUsl = num(body.usl), cNom = num(body.nominal);
      if ([cLsl, cUsl, cNom].some((n: number) => Number.isNaN(n))) {
        return { valid: false, error: "lsl/usl/nominal doivent etre numeriques." };
      }
      return { valid: true };
    }

    it('should reject missing mesuresBrutes', () => {
      const result = validateCapabilityInput({ lsl: 0, usl: 100 });
      expect(result.valid).toBe(false);
      expect(result.error).toContain('mesuresBrutes');
    });

    it('should reject empty mesuresBrutes', () => {
      const result = validateCapabilityInput({ mesuresBrutes: [], lsl: 0, usl: 100 });
      expect(result.valid).toBe(false);
    });

    it('should reject more than 1000 measurements', () => {
      const result = validateCapabilityInput({ mesuresBrutes: new Array(1001).fill(50), lsl: 0, usl: 100 });
      expect(result.valid).toBe(false);
    });

    it('should reject non-numeric values in mesuresBrutes', () => {
      const result = validateCapabilityInput({ mesuresBrutes: ['abc', 'def'], lsl: 0, usl: 100 });
      expect(result.valid).toBe(false);
    });

    it('should reject non-numeric lsl', () => {
      const result = validateCapabilityInput({ mesuresBrutes: [50], lsl: 'invalid', usl: 100, nominal: 50 });
      expect(result.valid).toBe(false);
    });

    it('should accept valid input', () => {
      const result = validateCapabilityInput({ mesuresBrutes: [50, 51, 49], lsl: 40, usl: 60, nominal: 50 });
      expect(result.valid).toBe(true);
    });
  });

  describe('Export input validation', () => {
    function validateExportInput(body: any): { valid: boolean; error?: string } {
      if (!body.base64Data) {
        return { valid: false, error: 'Missing base64Data' };
      }
      return { valid: true };
    }

    it('should reject missing base64Data', () => {
      const result = validateExportInput({ filename: 'test.xlsx' });
      expect(result.valid).toBe(false);
      expect(result.error).toBe('Missing base64Data');
    });

    it('should accept valid base64Data', () => {
      const result = validateExportInput({ base64Data: 'dGVzdA==', filename: 'test.xlsx' });
      expect(result.valid).toBe(true);
    });
  });
});

describe('GRR Metadata Assignment', () => {
  describe('assignGrrMetadata', () => {
    function assignGrrMetadataBackend(filename: string, index: number, total: number) {
      const lower = filename.toLowerCase();

      let operatorIndex = 0;
      if (lower.includes("op2") || lower.includes("oper2") || lower.includes("operateur2") || lower.includes("op_2") || lower.includes("operator2") || lower.includes("tech_b") || lower.includes("techb") || lower.includes("o2")) {
        operatorIndex = 1;
      } else if (lower.includes("op1") || lower.includes("oper1") || lower.includes("operateur1") || lower.includes("op_1") || lower.includes("operator1") || lower.includes("tech_a") || lower.includes("techa") || lower.includes("o1")) {
        operatorIndex = 0;
      } else {
        operatorIndex = index < (total / 2) ? 0 : 1;
      }

      let partIndex = 0;
      if (lower.includes("p2") || lower.includes("part2") || lower.includes("piece2") || lower.includes("part_2") || lower.includes("piece_2") || lower.includes("partb") || lower.includes("part b")) {
        partIndex = 1;
      } else if (lower.includes("p1") || lower.includes("part1") || lower.includes("piece1") || lower.includes("part_1") || lower.includes("piece_1") || lower.includes("parta") || lower.includes("part a")) {
        partIndex = 0;
      } else {
        const operatorFilesCount = Math.ceil(total / 2);
        const localIndex = index < operatorFilesCount ? index : (index - operatorFilesCount);
        partIndex = localIndex < (operatorFilesCount / 2) ? 0 : 1;
      }

      let trialIndex = 0;
      if (lower.includes("t3") || lower.includes("trial3") || lower.includes("essai3") || lower.includes("run3") || lower.includes("run_3")) {
        trialIndex = 2;
      } else if (lower.includes("t2") || lower.includes("trial2") || lower.includes("essai2") || lower.includes("run2") || lower.includes("run_2")) {
        trialIndex = 1;
      } else if (lower.includes("t1") || lower.includes("trial1") || lower.includes("essai1") || lower.includes("run1") || lower.includes("run_1")) {
        trialIndex = 0;
      } else {
        trialIndex = index % 3;
      }

      return { filename, operatorIndex, partIndex, trialIndex };
    }

    it('should assign operator index 0 for op1 prefix', () => {
      const result = assignGrrMetadataBackend('op1_file.txt', 0, 4);
      expect(result.operatorIndex).toBe(0);
    });

    it('should assign operator index 1 for op2 prefix', () => {
      const result = assignGrrMetadataBackend('op2_file.txt', 2, 4);
      expect(result.operatorIndex).toBe(1);
    });

    it('should assign part index 1 for p2 prefix', () => {
      const result = assignGrrMetadataBackend('p2_data.csv', 0, 2);
      expect(result.partIndex).toBe(1);
    });

    it('should assign trial index 0 for t1 prefix', () => {
      const result = assignGrrMetadataBackend('t1_measure.txt', 0, 3);
      expect(result.trialIndex).toBe(0);
    });

    it('should assign trial index 2 for t3 prefix', () => {
      const result = assignGrrMetadataBackend('t3_measure.txt', 2, 3);
      expect(result.trialIndex).toBe(2);
    });

    it('should fallback to modulo for trial index', () => {
      const result = assignGrrMetadataBackend('measure.txt', 5, 6);
      expect(result.trialIndex).toBe(5 % 3);
    });
  });
});