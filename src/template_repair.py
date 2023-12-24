with open(self.TEMPLATES_PATH, 'r') as f:
    json_data = dict(json.load(f))
    print(json_data)

    for name in json_data:
        name_data = dict(json_data[name])
        
        if 'Student' in name_data.keys():
            name_data['Students'] = name_data['Student']
            del name_data['Student']
            json_data[name] = name_data
            
    print(json_data)

    with open(self.TEMPLATES_PATH, 'w') as f:
        f.write(json.dumps(json_data))   
        f.close()   
        
    f.close()