class FileProcessor:
    @staticmethod
    def read_file(file_path):
        try:
            with open(file_path,'r') as f:
                return f.readlines()
        except Exception as e:
            print(f"文件读取失败:{file_path}-{str(e)}")
            return []
    @staticmethod
    def find_chunk(indicator,endflag,list):
        chunk = []
        i = 0
        find_target = False
        end_count = 0
        while i<len(list):
            if indicator in list[i]:
                find_target = True
                chunk.append(list[i])
                i += 1
                continue
            if find_target:
                if endflag == '**':
                    if endflag in list[i]:
                        end_count += 1
                    if end_count == 2:
                        break
                    else :
                        chunk.append(list[i])
                        i += 1
                        continue
                else:
                    if endflag in list[i]:
                        break
                    else :
                        chunk.append(list[i])
                        i += 1
                        continue
            i += 1
        return chunk
    @staticmethod
    def find_chunks(indicators,endflag,file_list):
        chunks = []
        for indicator in indicators.values():
            chunk = FileProcessor.find_chunk(indicator,endflag,file_list)
            chunks.append(chunk)
        return chunks